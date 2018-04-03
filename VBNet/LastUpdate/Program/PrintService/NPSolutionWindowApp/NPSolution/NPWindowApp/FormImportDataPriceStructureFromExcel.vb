Option Explicit On
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports Microsoft.Office.Interop


Public Class FormImportDataPriceStructureFromExcel
    Dim dt As DataTable
    Dim vList As ListViewItem
    Dim i As Integer
    Dim vPathFileName As String
    Dim vQuery As String
    Dim da As SqlDataAdapter
    Dim vExecuteQuery As SqlCommand
    Dim vdt As DataTable
    Dim vReadQuery As SqlDataReader
    Dim vCMD As SqlCommand

    Dim vCheckIsConfirm As Integer
    Dim vCheckItemExist As Integer
    Dim vCheckErrorFileExcel As Integer

    Dim frmPriceStructureRequest As FormPriceStructureRequest

    Private Sub GetDataHeaderFromExcel()
        'On Error Resume Next
        'Dim ex As New Excel.Application
        'Dim wb As Excel.Workbook
        'Dim ws As Excel.Worksheet

        ''Dim ex As New Microsoft.Office.Interop.Excel.Application
        ''Dim wb As Microsoft.Office.Interop.Excel.Workbook
        ''Dim ws As Microsoft.Office.Interop.Excel.Worksheet
        '' ---- �Դ excel worksheet.        
        'wb = ex.Workbooks.Open(vPathFileName)
        'ex.ActiveWorkbook.Sheets(1).Select()
        'ws = ex.ActiveWorkbook.ActiveSheet
        '' ---- ���ҧ DataTable        
        'Dim row As Integer
        'Dim HasData As Boolean = True
        'Dim dt As New DataTable("SheetOne")
        'Dim dr As DataRow
        'dt.Columns.Add("��¡��", GetType(String))
        'dt.Columns.Add("��ҷ���˹�", GetType(String))

        'row = 1 ' ---- ���������������Ƿ�� 2 (���á��� 1, ��������á��� 1)        
        'Do
        '    dr = dt.NewRow
        '    dr("��¡��") = ws.Cells(row, 1).value
        '    dr("��ҷ���˹�") = ws.Cells(row, 2).value
        '    dt.Rows.Add(dr)
        '    row += 1
        '    If row > 3 Then
        '        HasData = False
        '    End If
        'Loop While (HasData)

        'DGHeader.DataSource = dt

        'DGHeader.Columns(0).Width = 300
        'DGHeader.Columns(1).Width = 650
        'wb.Close()
    End Sub
    Private Sub GetDataDetailsFromExcel()
        'On Error Resume Next
        'Dim ex As New Excel.Application
        'Dim wb As Excel.Workbook
        'Dim ws As Excel.Worksheet
        ''Dim ex As New Microsoft.Office.Interop.Excel.Application
        ''Dim wb As Microsoft.Office.Interop.Excel.Workbook
        ''Dim ws As Microsoft.Office.Interop.Excel.Worksheet
        'Dim i As Integer
        'Dim n As Integer
        'Dim vItemCode As String
        'Dim vUnitCode As String
        'Dim vCheckItemCode As String


        '' ---- �Դ excel worksheet.        
        'wb = ex.Workbooks.Open(vPathFileName)
        'ex.ActiveWorkbook.Sheets(1).Select()
        'ws = ex.ActiveWorkbook.ActiveSheet
        '' ---- ���ҧ DataTable        
        'Dim row As Integer
        'Dim HasData As Boolean = True
        'Dim HasDataExcel As Boolean = True
        'Dim dt As New DataTable("SheetOne")
        'Dim dr As DataRow
        'dt.Columns.Add("�����Թ���", GetType(String))
        'dt.Columns.Add("�����Թ���", GetType(String))
        'dt.Columns.Add("˹��¹Ѻ���", GetType(String))
        'dt.Columns.Add("D/O", GetType(String))
        'dt.Columns.Add("��ǹŴ˹�Һ��", GetType(String))
        'dt.Columns.Add("�ع�ѭ��", GetType(String))
        'dt.Columns.Add("��ǹŴ���1", GetType(String))
        'dt.Columns.Add("��Ť����ǹŴ���1", GetType(Double))
        'dt.Columns.Add("��ǹŴ���2", GetType(String))
        'dt.Columns.Add("��Ť����ǹŴ���2", GetType(String))
        'dt.Columns.Add("��ǹŴ���3", GetType(String))
        'dt.Columns.Add("��Ť����ǹŴ���3", GetType(String))
        'dt.Columns.Add("��ǹŴ���4", GetType(String))
        'dt.Columns.Add("��Ť����ǹŴ���4", GetType(String))
        'dt.Columns.Add("��ǹŴRebate", GetType(String))
        'dt.Columns.Add("��Ť����ǹŴRebate", GetType(String))
        'dt.Columns.Add("��ǹŴ�����", GetType(String))
        'dt.Columns.Add("�ع�ط��", GetType(String))
        'dt.Columns.Add("���Ҵ�ع", GetType(String))
        'dt.Columns.Add("��Ť�ҧ��Ҵ�ع", GetType(String))
        'dt.Columns.Add("��Ң������", GetType(String))
        'dt.Columns.Add("��Ң�������١���", GetType(String))
        'dt.Columns.Add("�ɳ�", GetType(String))
        'dt.Columns.Add("��Ť���ɳ�", GetType(String))
        'dt.Columns.Add("�ع���� 7%", GetType(String))
        'dt.Columns.Add("����ç��ҵԴ���", GetType(String))
        'dt.Columns.Add("��ԡ��", GetType(String))
        'dt.Columns.Add("�ع��Ҵ", GetType(String))
        'dt.Columns.Add("�͡����ʵ�͡ (%)", GetType(String))
        'dt.Columns.Add("�͡����ʵ�͡ (�ҷ)", GetType(String))
        'dt.Columns.Add("���â�»�ա (%)", GetType(String))
        'dt.Columns.Add("���â�»�ա (�ҷ)", GetType(String))
        'dt.Columns.Add("���â���� (%)", GetType(String))
        'dt.Columns.Add("���â���� (�ҷ)", GetType(String))
        'dt.Columns.Add("SmartPoint(%)", GetType(String))
        'dt.Columns.Add("SmartPoint(�ҷ)", GetType(String))
        'dt.Columns.Add("���", GetType(String))
        'dt.Columns.Add("�ͧ��", GetType(String))
        'dt.Columns.Add("����Ԫ��� (%)", GetType(String))
        'dt.Columns.Add("����Ԫ��� (�ҷ)", GetType(String))
        'dt.Columns.Add("�Ҥ����", GetType(String))
        'dt.Columns.Add("�Ҥ�1(ʴ�Ѻ�ͧ)", GetType(String))
        'dt.Columns.Add("�Ҥ�1(ʴ�����)", GetType(String))
        'dt.Columns.Add("�Ҥ�1(�����Ѻ�ͧ)", GetType(String))
        'dt.Columns.Add("�Ҥ�1(���������)", GetType(String))
        'dt.Columns.Add("�Ҥ�2", GetType(String))
        'dt.Columns.Add("�ѹ����Ѻ�Ҥ�", GetType(String))
        'dt.Columns.Add("�����˵�", GetType(String))
        'dt.Columns.Add("���â�鹵�", GetType(String))

        'Dim vCheckPrice1 As Double
        'Dim vCheckPrice2 As Double
        'Dim vCheckPrice3 As Double
        'Dim vCheckPrice4 As Double
        'Dim vCheckPrice5 As Double

        'i = 0
        'row = 9
        'Do
        '    i = i + 1
        '    row += 1
        '    vCheckItemCode = ws.Cells(row, 1).value
        '    'MsgBox(ws.Cells(row, 1).value, MsgBoxStyle.Critical, "")
        '    'If IsDBNull(ws.Cells(row, 1).value) And (ws.Cells(row, 1).value) = 0 Then
        '    If IsDBNull(vCheckItemCode) Or vCheckItemCode = "" Then
        '        HasDataExcel = False
        '    End If
        '    Me.PB101.Maximum = i
        '    n = 1
        'Loop While (HasDataExcel)
        'row = 9 ' ---- ���������������Ƿ�� 2 (���á��� 1, ��������á��� 1)        
        '    Do
        '    dr = dt.NewRow
        '    dr("�����Թ���") = ws.Cells(row, 1).value
        '    Dim vCountItem As Integer
        '    If row > 9 Then
        '        For vCountItem = 9 To row - 1
        '            If ws.Cells(row, 1).value = ws.Cells(vCountItem, 1).value Then
        '                MsgBox("�����Թ��� " & ws.Cells(row, 1).value & " 㹺�÷Ѵ��� " & row & " ��ӡѺ�����Թ��Һ�÷Ѵ��� " & vCountItem & " ")
        '                Me.PB101.Value = 0
        '                Me.DGHeader.DataSource = Nothing
        '                Me.DGDetails.DataSource = Nothing
        '                Exit Sub
        '            End If
        '        Next
        '    End If
        '    dr("�����Թ���") = ws.Cells(row, 2).value
        '    dr("˹��¹Ѻ���") = ws.Cells(row, 3).value
        '    If ws.Cells(row, 3).value = "" Then
        '        MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ������˹� ˹��¹Ѻ �������ö�ѹ�֡��������")
        '        Me.PB101.Value = 0
        '        Me.DGHeader.DataSource = Nothing
        '        Me.DGDetails.DataSource = Nothing
        '        Exit Sub
        '    End If
        '    dr("D/O") = ws.Cells(row, 4).value
        '    dr("��ǹŴ˹�Һ��") = ws.Cells(row, 5).value
        '    dr("�ع�ѭ��") = ws.Cells(row, 6).value
        '    dr("��ǹŴ���1") = ws.Cells(row, 7).value
        '    dr("��Ť����ǹŴ���1") = ws.Cells(row, 8).value
        '    dr("��ǹŴ���2") = ws.Cells(row, 9).value
        '    dr("��Ť����ǹŴ���2") = ws.Cells(row, 10).value
        '    dr("��ǹŴ���3") = ws.Cells(row, 11).value
        '    dr("��Ť����ǹŴ���3") = ws.Cells(row, 12).value
        '    dr("��ǹŴ���4") = ws.Cells(row, 13).value
        '    dr("��Ť����ǹŴ���4") = ws.Cells(row, 14).value
        '    dr("��ǹŴRebate") = ws.Cells(row, 15).value
        '    dr("��Ť����ǹŴRebate") = ws.Cells(row, 16).value
        '    dr("��ǹŴ�����") = ws.Cells(row, 17).value
        '    dr("�ع�ط��") = ws.Cells(row, 18).value
        '    dr("���Ҵ�ع") = ws.Cells(row, 19).value
        '    dr("��Ť�ҧ��Ҵ�ع") = ws.Cells(row, 20).value
        '    dr("��Ң������") = ws.Cells(row, 21).value
        '    dr("��Ң�������١���") = ws.Cells(row, 22).value
        '    dr("�ɳ�") = ws.Cells(row, 23).value
        '    dr("��Ť���ɳ�") = ws.Cells(row, 24).value
        '    dr("�ع���� 7%") = ws.Cells(row, 25).value
        '    dr("����ç��ҵԴ���") = ws.Cells(row, 26).value
        '    dr("��ԡ��") = ws.Cells(row, 27).value
        '    dr("�ع��Ҵ") = ws.Cells(row, 28).value
        '    dr("�͡����ʵ�͡ (%)") = ws.Cells(row, 29).value
        '    dr("�͡����ʵ�͡ (�ҷ)") = ws.Cells(row, 30).value
        '    dr("���â�»�ա (%)") = ws.Cells(row, 31).value
        '    dr("���â�»�ա (�ҷ)") = ws.Cells(row, 32).value
        '    dr("���â���� (%)") = ws.Cells(row, 33).value
        '    dr("���â���� (�ҷ)") = ws.Cells(row, 34).value
        '    dr("SmartPoint(%)") = ws.Cells(row, 35).value
        '    dr("SmartPoint(�ҷ)") = ws.Cells(row, 36).value
        '    dr("���") = ws.Cells(row, 37).value
        '    dr("�ͧ��") = ws.Cells(row, 38).value
        '    dr("����Ԫ��� (%)") = ws.Cells(row, 39).value
        '    dr("����Ԫ��� (�ҷ)") = ws.Cells(row, 40).value
        '    dr("�Ҥ����") = ws.Cells(row, 41).value
        '    dr("�Ҥ�1(ʴ�Ѻ�ͧ)") = ws.Cells(row, 42).value
        '    vCheckPrice1 = ws.Cells(row, 42).value
        '    If vCheckPrice1 = 0 Then
        '        MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ������˹� �Ҥ�1(ʴ�Ѻ�ͧ) �������ö�ѹ�֡��������")
        '        Me.PB101.Value = 0
        '        Me.DGHeader.DataSource = Nothing
        '        Me.DGDetails.DataSource = Nothing
        '        Exit Sub
        '    End If
        '    dr("�Ҥ�1(ʴ�����)") = ws.Cells(row, 43).value
        '    vCheckPrice2 = ws.Cells(row, 43).value
        '    If vCheckPrice2 = 0 Then
        '        MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ������˹� �Ҥ�1(ʴ�����) �������ö�ѹ�֡��������")
        '        Me.PB101.Value = 0
        '        Me.DGHeader.DataSource = Nothing
        '        Me.DGDetails.DataSource = Nothing
        '        Exit Sub
        '    End If
        '    dr("�Ҥ�1(�����Ѻ�ͧ)") = ws.Cells(row, 44).value
        '    vCheckPrice3 = ws.Cells(row, 44).value
        '    If vCheckPrice3 = 0 Then
        '        MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ������˹� �Ҥ�1(�����Ѻ�ͧ) �������ö�ѹ�֡��������")
        '        Me.PB101.Value = 0
        '        Me.DGHeader.DataSource = Nothing
        '        Me.DGDetails.DataSource = Nothing
        '        Exit Sub
        '    End If
        '    dr("�Ҥ�1(���������)") = ws.Cells(row, 45).value
        '    vCheckPrice4 = ws.Cells(row, 45).value
        '    If vCheckPrice4 = 0 Then
        '        MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ������˹� �Ҥ�1(���������) �������ö�ѹ�֡��������")
        '        Me.PB101.Value = 0
        '        Me.DGHeader.DataSource = Nothing
        '        Me.DGDetails.DataSource = Nothing
        '        Exit Sub
        '    End If
        '    dr("�Ҥ�2") = ws.Cells(row, 46).value
        '    vCheckPrice5 = ws.Cells(row, 46).value
        '    If vCheckPrice5 = 0 Then
        '        MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ������˹��Ҥҷ�� 2 �������ö�ѹ�֡��������")
        '        Me.PB101.Value = 0
        '        Me.DGHeader.DataSource = Nothing
        '        Me.DGDetails.DataSource = Nothing
        '        Exit Sub
        '    End If

        '    Dim vUpDateDay As Date
        '    Dim vUpDateDay1 As String

        '    vUpDateDay = ws.Cells(row, 47).value
        '    vUpDateDay1 = vUpDateDay.Day & "/" & vUpDateDay.Month & "/" & vUpDateDay.Year

        '    dr("�ѹ����Ѻ�Ҥ�") = vUpDateDay1

        '    If InStr(Microsoft.VisualBasic.Right(ws.Cells(row, 47).value, 4), "/") = 0 Then
        '        'MsgBox(Microsoft.VisualBasic.Mid(Microsoft.VisualBasic.Right(ws.Cells(row, 47).value, 4), 2, 1))
        '        If Microsoft.VisualBasic.Mid(Microsoft.VisualBasic.Right(ws.Cells(row, 47).value, 4), 2, 1) <> "0" Then
        '            MsgBox("�Թ��� ���� " & ws.Cells(row, 1).value & " " & ws.Cells(row, 2).value & " ��˹��ѹ����Ѻ�Ҥ��繾ط��ѡ�Ҫ �������ö�ѹ�֡��������")
        '            Me.PB101.Value = 0
        '            Me.DGHeader.DataSource = Nothing
        '            Me.DGDetails.DataSource = Nothing
        '            Exit Sub
        '        End If
        '    End If

        '    dr("�����˵�") = ws.Cells(row, 48).value
        '    dr("���â�鹵�") = ws.Cells(row, 49).value

        '    vItemCode = ws.Cells(row, 1).value
        '    vUnitCode = ws.Cells(row, 3).value

        '    dt.Rows.Add(dr)

        '    vQuery = "select isnull(count(itemcode),0) as vCount from dbo.bcpricelist where itemcode = '" & vItemCode & "' and unitcode = '" & vUnitCode & "'"
        '    vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '    vReadQuery = vExecuteQuery.ExecuteReader
        '    While vReadQuery.Read
        '        vCheckItemExist = vReadQuery(0)
        '    End While
        '    vReadQuery.Close()

        '    If vCheckItemExist = 0 Then
        '        vCheckErrorFileExcel = 1
        '        MsgBox("�����Թ��� " & vItemCode & " ����բ�����������к� �������ö���ç���ҧ�Ҥ��� ��س�����͡�����е�Ǩ�ͺ�����Ţͧ�Թ��Ҵѧ����Ǵ���", MsgBoxStyle.Critical, "Send Error Message")
        '    End If

        '    row += 1
        '    If IsDBNull(ws.Cells(row, 1).value) Or ws.Cells(row, 1).value = "" Then
        '        HasData = False
        '    End If
        '    Me.PB101.Value = n
        '    n = n + 1
        'Loop While (HasData)
        'DGDetails.DataSource = dt
        'DGDetails.Columns(0).Width = 150
        'DGDetails.Columns(1).Width = 300   
        'wb.Close()
    End Sub

    Private Sub FormImportDataPriceStructureFromExcel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        Me.PBNew.Visible = True
        Me.PBConfirm.Visible = False
    End Sub

    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DGDetails.RowPostPaint
        Using b As SolidBrush = New SolidBrush(Me.DGDetails.RowHeadersDefaultCellStyle.ForeColor)

            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), _
                                    Me.DGDetails.DefaultCellStyle.Font, _
                                    b, e.RowBounds.Location.X + 5, _
                                    e.RowBounds.Location.Y + 5)
        End Using
        '��������Ţ��� GridView
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocno As String
        Dim vDocDate As String
        Dim vTargetProfit As String
        Dim vProfit As String
        Dim vSmartPoint As String
        Dim vMemberDiscount As String
        Dim vFileDataSource As String
        Dim vPathFile As String
        '------------------------------------------
        Dim vItemCode As String
        Dim vItemName As String
        Dim vSaleUnitCode As String
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vDiscountBillWord As String
        Dim vDiscountBill1 As Double
        Dim vDiscountBillAmount As Double
        Dim vAccCost As Double
        Dim vDiscountFollow1Word As String
        Dim vDiscountFollow11 As Double
        Dim vDiscountFollow1Amount As Double
        Dim vDiscountFollow1After As Double
        Dim vDiscountFollow2Word As String
        Dim vDiscountFollow21 As Double
        Dim vDiscountFollow2Amount As Double
        Dim vDiscountFollow2After As Double
        Dim vDiscountFollow3Word As String
        Dim vDiscountFollow31 As Double
        Dim vDiscountFollow3Amount As Double
        Dim vDiscountFollow3After As Double
        Dim vDiscountFollow4Word As String
        Dim vDiscountFollow41 As Double
        Dim vDiscountFollow4Amount As Double
        Dim vDiscountFollow4After As Double
        Dim vDiscountRebateWord As String
        Dim vDiscountRebate1 As Double
        Dim vDiscountRebateAmount As Double
        Dim vDiscountRebateAfter As Double
        Dim vDiscountSpecialWord As String
        Dim vDiscountSpecial1 As Double
        Dim vDiscountSpecialAmount As Double
        Dim vNetCost As Double
        Dim vLossBudgetWord As String
        Dim vLossBudget1 As Double
        Dim vLossBudgetAmount As Double
        Dim vLossBudgetAfter As Double
        Dim vTransferInWord As String
        Dim vTransferIn1 As Double
        Dim vTransferOutWord As String
        Dim vTransferOut1 As Double
        Dim vAdvertiseWord As String
        Dim vAdvertise1 As Double
        Dim vAdvertiseAmount As Double
        Dim vAdvertiseAfter As Double
        Dim vVatCost As Double
        Dim vVatAmount As Double
        Dim vSetupWord As String
        Dim vSetupAmount As Double
        Dim vServiceWord As String
        Dim vServiceAmount As Double
        Dim vMarketCost As Double
        Dim vPointWord As String
        Dim vPoint1 As Double
        Dim vPointAmount As Double
        Dim vPointAfter As Double
        Dim vTargetWord As String
        Dim vTargetAmount As Double
        Dim vPremiumWord As String
        Dim vPremiumAmount As Double
        Dim vCommissionWord As String
        Dim vCommission1 As Double
        Dim vCommissionAmount As Double
        Dim vCommissionAfter As Double
        Dim vGrossProfitPercent As String
        Dim vGrossProfitAmount As Double
        Dim vInterestStockPercent As String
        Dim vInterestStockAmount As Double
        Dim vProfitPercent As String
        Dim vProfitAmount As Double
        Dim vProfitPercent_W As String
        Dim vProfitAmount_W As Double
        Dim vMyDescription As String
        Dim vMyDescriptionSub As String
        Dim vTransferInAfter As Double
        Dim vTransferOutAfter As Double
        Dim vVatWord As String
        Dim vSetupAfter As Double
        Dim vTargetAfter As Double
        Dim vPremiumAfter As Double
        '--------------------------------------------------------------

        '---------------------------------------------------------------
        Dim i As Integer
        Dim vLenFilePath As Integer
        Dim vLenMark As Integer

        If vCheckErrorFileExcel = 0 Then
            If vCheckIsConfirm = 0 Then
                If Me.DGHeader.Rows.Count > 0 And Me.DGDetails.Rows.Count > 0 Then
                    Me.PB101.Value = 1
                    If Me.TextDocNo.Text = "" Then
                        vQuery = "exec dbo.USP_PS_NewDocno"
                        da = New SqlDataAdapter(vQuery, vConnection)
                        ds = New DataSet
                        da.Fill(ds, "NewDocno")
                        vdt = ds.Tables("NewDocno")
                        vDocno = vdt.Rows(0).Item("newdocno")
                        vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
                    Else
                        vDocno = Me.TextDocNo.Text
                        vDocDate = Me.DocDate.Text
                    End If


                    Try
                        vQuery = "begin tran"
                        vExecuteQuery = New SqlCommand(vQuery, vConnection)
                        vExecuteQuery.ExecuteNonQuery()
                        If Not IsDBNull(Me.DGHeader.Rows(0).Cells(1).Value) Then
                            vTargetProfit = CType((CType(Me.DGHeader.Rows(0).Cells(1).Value, Double) * 100), String) & "%"
                        Else
                            vTargetProfit = 0
                        End If
                        If Not IsDBNull(Me.DGHeader.Rows(1).Cells(1).Value) Then
                            vProfit = CType((CType(Me.DGHeader.Rows(1).Cells(1).Value, Double) * 100), String) & "%"
                        Else
                            vProfit = 0
                        End If
                        If Not IsDBNull(Me.DGHeader.Rows(2).Cells(1).Value) Then
                            vSmartPoint = CType((CType(Me.DGHeader.Rows(2).Cells(1).Value, Double) * 100), String) & "%"
                        Else
                            vSmartPoint = 0
                        End If
                        If Not IsDBNull(Me.DGHeader.Rows(3).Cells(1).Value) Then
                            vMemberDiscount = CType((CType(Me.DGHeader.Rows(3).Cells(1).Value, Double) * 100), String) & "%"
                        Else
                            vMemberDiscount = 0
                        End If
                        vMyDescription = Me.TextMyDescription.Text
                        vPathFile = Trim(Me.LBLFileName.Text)
                        vLenMark = InStr(Me.LBLFileName.Text, "@")
                        vLenFilePath = Len(Me.LBLFileName.Text)
                        vFileDataSource = Microsoft.VisualBasic.Right(Me.LBLFileName.Text, vLenFilePath - vLenMark)
                        vQuery = "exec dbo.USP_PS_InsertPriceStructureSet1 '" & vDocno & "','" & vDocDate & "','" & vTargetProfit & "','" & vProfit & "','" & vSmartPoint & "','" & vMemberDiscount & "','" & vFileDataSource & "','" & vPathFile & "','" & vMyDescription & "'"
                        vExecuteQuery = New SqlCommand(vQuery, vConnection)
                        vExecuteQuery.ExecuteNonQuery()

                        Me.PB101.Maximum = Me.DGDetails.RowCount - 2

                        For i = 0 To Me.DGDetails.RowCount - 2

                            vItemCode = Trim(Me.DGDetails.Rows(i).Cells(0).Value)
                            vItemName = Trim(Me.DGDetails.Rows(i).Cells(1).Value)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(2).Value) Then
                                vSaleUnitCode = Trim(Me.DGDetails.Rows(i).Cells(2).Value)
                            Else
                                vSaleUnitCode = ""
                            End If
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(3).Value) Then
                                vDO = Trim(Me.DGDetails.Rows(i).Cells(3).Value)
                            Else
                                vDO = 0
                            End If
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(42).Value) Then
                                vPriceSet = Me.DGDetails.Rows(i).Cells(42).Value
                            Else
                                vPriceSet = 0
                            End If
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(4).Value) Then
                                vDiscountBillWord = CType((CType(Me.DGDetails.Rows(i).Cells(4).Value, Double) * 100), String) & "%"
                                vDiscountBill1 = (CType(Me.DGDetails.Rows(i).Cells(4).Value, Double) * 100)
                            Else
                                vDiscountBillWord = ""
                                vDiscountBill1 = 0
                            End If
                            vDiscountBillAmount = CalcAmount(vDO, vDiscountBill1)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(5).Value) Then
                                vAccCost = Me.DGDetails.Rows(i).Cells(5).Value
                            Else
                                vAccCost = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(6).Value) Then
                                vDiscountFollow1Word = CType((CType(Me.DGDetails.Rows(i).Cells(6).Value, Double) * 100), String) & "%"
                                vDiscountFollow11 = (CType(Me.DGDetails.Rows(i).Cells(6).Value, Double) * 100)
                            Else
                                vDiscountFollow1Word = ""
                                vDiscountFollow11 = 0
                            End If
                            vDiscountFollow1Amount = CalcAmount(vAccCost, vDiscountFollow11)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(7).Value) Then
                                vDiscountFollow1After = Me.DGDetails.Rows(i).Cells(7).Value
                            Else
                                vDiscountFollow1After = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(8).Value) Then
                                vDiscountFollow2Word = CType((CType(Me.DGDetails.Rows(i).Cells(8).Value, Double) * 100), String) & "%"
                                vDiscountFollow21 = (CType(Me.DGDetails.Rows(i).Cells(8).Value, Double) * 100)
                            Else
                                vDiscountFollow2Word = ""
                                vDiscountFollow21 = 0
                            End If
                            vDiscountFollow2Amount = CalcAmount(vDiscountFollow1After, vDiscountFollow21)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(9).Value) Then
                                vDiscountFollow2After = Me.DGDetails.Rows(i).Cells(9).Value
                            Else
                                vDiscountFollow2After = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(10).Value) Then
                                vDiscountFollow3Word = CType((CType(Me.DGDetails.Rows(i).Cells(10).Value, Double) * 100), String) & "%"
                                vDiscountFollow31 = (CType(Me.DGDetails.Rows(i).Cells(10).Value, Double) * 100)
                            Else
                                vDiscountFollow3Word = ""
                                vDiscountFollow31 = 0
                            End If
                            vDiscountFollow3Amount = CalcAmount(vDiscountFollow2After, vDiscountFollow31)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(11).Value) Then
                                vDiscountFollow3After = Me.DGDetails.Rows(i).Cells(11).Value
                            Else
                                vDiscountFollow3After = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(12).Value) Then
                                vDiscountFollow4Word = CType((CType(Me.DGDetails.Rows(i).Cells(12).Value, Double) * 100), String) & "%"
                                vDiscountFollow41 = (CType(Me.DGDetails.Rows(i).Cells(12).Value, Double) * 100)
                            Else
                                vDiscountFollow4Word = ""
                                vDiscountFollow41 = 0
                            End If
                            vDiscountFollow4Amount = CalcAmount(vDiscountFollow3After, vDiscountFollow41)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(13).Value) Then
                                vDiscountFollow4After = Me.DGDetails.Rows(i).Cells(13).Value
                            Else
                                vDiscountFollow4After = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(14).Value) Then
                                vDiscountRebateWord = CType((CType(Me.DGDetails.Rows(i).Cells(14).Value, Double) * 100), String) & "%"
                                vDiscountRebate1 = (CType(Me.DGDetails.Rows(i).Cells(14).Value, Double) * 100)
                            Else
                                vDiscountRebateWord = ""
                                vDiscountRebate1 = 0
                            End If
                            vDiscountRebateAmount = CalcAmount(vDiscountFollow4After, vDiscountRebate1)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(15).Value) Then
                                vDiscountRebateAfter = Me.DGDetails.Rows(i).Cells(15).Value
                            Else
                                vDiscountRebateAfter = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(16).Value) Then
                                vDiscountSpecialWord = Me.DGDetails.Rows(i).Cells(16).Value
                                vDiscountSpecial1 = Me.DGDetails.Rows(i).Cells(16).Value
                            Else
                                vDiscountSpecialWord = 0
                                vDiscountSpecial1 = 0
                            End If
                            vDiscountSpecialAmount = vDiscountSpecial1
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(17).Value) Then
                                vNetCost = Me.DGDetails.Rows(i).Cells(17).Value
                            Else
                                vNetCost = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(18).Value) Then
                                vLossBudgetWord = CType((CType(Me.DGDetails.Rows(i).Cells(18).Value, Double) * 100), String) & "%"
                                vLossBudget1 = (CType(Me.DGDetails.Rows(i).Cells(18).Value, Double) * 100)
                            Else
                                vLossBudgetWord = ""
                                vLossBudget1 = 0
                            End If
                            vLossBudgetAmount = CalcAmount(vNetCost, vLossBudget1)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(19).Value) Then
                                vLossBudgetAfter = Me.DGDetails.Rows(i).Cells(19).Value
                            Else
                                vLossBudgetAfter = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(20).Value) Then
                                vTransferInWord = Me.DGDetails.Rows(i).Cells(20).Value
                                vTransferIn1 = Me.DGDetails.Rows(i).Cells(20).Value
                                vTransferInAfter = vLossBudgetAfter + vTransferIn1
                            Else
                                vTransferIn1 = 0
                                vTransferInWord = ""
                                vTransferInAfter = vLossBudgetAfter + vTransferIn1
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(21).Value) Then
                                vTransferOutWord = Me.DGDetails.Rows(i).Cells(21).Value
                                vTransferOut1 = Me.DGDetails.Rows(i).Cells(21).Value
                                vTransferOutAfter = vTransferInAfter + vTransferOut1
                            Else
                                vTransferOutWord = ""
                                vTransferOut1 = 0
                                vTransferOutAfter = vTransferInAfter + vTransferOut1
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(22).Value) Then
                                vAdvertiseWord = CType((CType(Me.DGDetails.Rows(i).Cells(22).Value, Double) * 100), String) & "%"
                                vAdvertise1 = (CType(Me.DGDetails.Rows(i).Cells(22).Value, Double) * 100)
                            Else
                                vAdvertiseWord = ""
                                vAdvertise1 = 0
                            End If
                            vAdvertiseAmount = CalcAmount(vTransferOutAfter, vAdvertise1)
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(23).Value) Then
                                vAdvertiseAfter = Me.DGDetails.Rows(i).Cells(23).Value
                            Else
                                vAdvertiseAfter = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(24).Value) Then
                                vVatCost = Me.DGDetails.Rows(i).Cells(24).Value
                                vVatAmount = (vAdvertiseAfter * 7) / 100
                                vVatWord = "7%"
                            Else
                                vVatCost = 0
                                vVatAmount = 0
                                vVatWord = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(25).Value) Then
                                vSetupWord = Me.DGDetails.Rows(i).Cells(25).Value
                                vSetupAmount = Me.DGDetails.Rows(i).Cells(25).Value
                                vSetupAfter = vVatCost + vSetupAmount
                            Else
                                vSetupWord = ""
                                vSetupAmount = 0
                                vSetupAfter = vVatCost + vSetupAmount
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(26).Value) Then
                                vServiceWord = Me.DGDetails.Rows(i).Cells(26).Value
                                vServiceAmount = Me.DGDetails.Rows(i).Cells(26).Value
                            Else
                                vServiceWord = ""
                                vServiceAmount = 0
                            End If
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(27).Value) Then
                                vMarketCost = Me.DGDetails.Rows(i).Cells(27).Value
                            Else
                                vMarketCost = 0
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(34).Value) Then
                                vPointWord = CType((CType(Me.DGDetails.Rows(i).Cells(34).Value, Double) * 100), String) & "%"
                            Else
                                vPointWord = ""
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(35).Value) Then
                                vPoint1 = Me.DGDetails.Rows(i).Cells(35).Value
                            Else
                                vPoint1 = 0
                            End If
                            vPointAmount = vPoint1
                            vPointAfter = vMarketCost + vPointAmount

                            If vPointWord = "" And vPointAmount = 0 Then
                                MsgBox("�����Թ��� " & vItemCode & "   " & vItemName & " ����դ�Ңͧ Smart Point �������ö�ѹ�֡��������  ��س���䢢����š�͹�ѹ�֡����")
                                vQuery = "rollback tran"
                                vExecuteQuery = New SqlCommand(vQuery, vConnection)
                                vExecuteQuery.ExecuteNonQuery()
                                Exit Sub
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(36).Value) Then
                                vTargetWord = Me.DGDetails.Rows(i).Cells(36).Value
                                vTargetAmount = Me.DGDetails.Rows(i).Cells(36).Value
                                vTargetAfter = vPointAfter + vTargetAmount
                            Else
                                vTargetWord = ""
                                vTargetAmount = 0
                                vTargetAfter = vPointAfter + vTargetAmount
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(37).Value) Then
                                vPremiumWord = Me.DGDetails.Rows(i).Cells(37).Value
                                vPremiumAmount = Me.DGDetails.Rows(i).Cells(37).Value
                                vPremiumAfter = vTargetAfter + vPremiumAmount
                            Else
                                vPremiumWord = ""
                                vPremiumAmount = 0
                                vPremiumAfter = vTargetAfter + vPremiumAmount
                            End If

                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(38).Value) Then
                                vCommissionWord = CType((CType(Me.DGDetails.Rows(i).Cells(38).Value, Double) * 100), String) & "%"
                            Else
                                vCommissionWord = ""
                            End If
                            vCommissionAmount = vCommission1
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(39).Value) Then
                                vCommission1 = Me.DGDetails.Rows(i).Cells(39).Value
                            Else
                                vCommission1 = 0
                            End If
                            vCommissionAmount = vCommission1
                            vCommissionAfter = vPremiumAfter + vCommissionAmount

                            vGrossProfitPercent = (Me.DGDetails.Rows(i).Cells(48).Value * 100)
                            vGrossProfitAmount = ((Me.DGDetails.Rows(i).Cells(48).Value * 100) * vMarketCost) / 100
                            vInterestStockPercent = (Me.DGDetails.Rows(i).Cells(28).Value * 100)
                            vInterestStockAmount = Me.DGDetails.Rows(i).Cells(29).Value
                            vProfitPercent = (Me.DGDetails.Rows(i).Cells(30).Value * 100)
                            vProfitAmount = Me.DGDetails.Rows(i).Cells(31).Value
                            vProfitPercent_W = (Me.DGDetails.Rows(i).Cells(32).Value * 100)
                            vProfitAmount_W = Me.DGDetails.Rows(i).Cells(33).Value
                            If Not IsDBNull(Me.DGDetails.Rows(i).Cells(47).Value) Then
                                vMyDescriptionSub = Me.DGDetails.Rows(i).Cells(47).Value
                            Else
                                vMyDescriptionSub = ""
                            End If
                            '---------------------------------------------------------------------------------------------------
                            Dim vFromQTY As Double
                            Dim vToQTY As Double
                            Dim vPriceSet2 As Double
                            Dim vIsPriceUpdate As Integer = 1
                            Dim vToUpdateDate As String = Me.DGDetails.Rows(i).Cells(46).Value
                            Dim vIsUpdate As Integer = 0
                            Dim vIsPrintLabel As Integer = 0
                            Dim vPrice1CashRec As Double
                            Dim vPrice1CashDel As Double
                            Dim vPrice1CreditRec As Double
                            Dim vPrice1CreditDel As Double
                            '---------------------------------------------------------------------------------------------------

                            vSaleUnitCode = Me.DGDetails.Rows(i).Cells(2).Value
                            vFromQTY = 1
                            vToQTY = 99999
                            vPrice1CashRec = Me.DGDetails.Rows(i).Cells(41).Value
                            vPrice1CashDel = Me.DGDetails.Rows(i).Cells(42).Value
                            vPrice1CreditRec = Me.DGDetails.Rows(i).Cells(43).Value
                            vPrice1CreditDel = Me.DGDetails.Rows(i).Cells(44).Value
                            vPriceSet2 = Me.DGDetails.Rows(i).Cells(45).Value


                            vQuery = "exec dbo.USP_PS_InsertPriceStructureSubSet '" & vDocno & "','" & vItemCode & "','" & vItemName & "','" & vSaleUnitCode & "'," & vDO & ", " _
                            & "" & vPriceSet & ",'" & vDiscountBillWord & "'," & vDiscountBillAmount & "," & vAccCost & "," _
                            & "'" & vDiscountFollow1Word & "'," & vDiscountFollow1Amount & "," & vDiscountFollow1After & "," _
                            & "'" & vDiscountFollow2Word & "'," & vDiscountFollow2Amount & "," & vDiscountFollow2After & "," _
                            & "'" & vDiscountFollow3Word & "'," & vDiscountFollow3Amount & "," & vDiscountFollow3After & "," _
                            & "'" & vDiscountFollow4Word & "'," & vDiscountFollow4Amount & "," & vDiscountFollow4After & "," _
                            & "'" & vDiscountRebateWord & "'," & vDiscountRebateAmount & "," & vDiscountRebateAfter & "," _
                            & "'" & vDiscountSpecialWord & "'," & vDiscountSpecialAmount & "," & vNetCost & "," _
                            & "'" & vLossBudgetWord & "'," & vLossBudgetAmount & "," & vLossBudgetAfter & "," _
                            & "'" & vTransferInWord & "'," & vTransferIn1 & "," & vTransferInAfter & "," _
                            & "'" & vTransferOutWord & "'," & vTransferOut1 & "," & vTransferOutAfter & "," _
                            & "'" & vAdvertiseWord & "'," & vAdvertiseAmount & "," & vAdvertiseAfter & "," _
                            & "'" & vVatWord & "'," & vVatCost & "," & vVatAmount & "," _
                            & "'" & vSetupWord & "'," & vSetupAmount & "," & vSetupAfter & "," _
                            & "'" & vServiceWord & "'," & vServiceAmount & "," & vMarketCost & "," _
                            & "'" & vPointWord & "'," & vPointAmount & "," & vPointAfter & "," _
                            & "'" & vTargetWord & "'," & vTargetAmount & "," & vTargetAfter & "," _
                            & "'" & vPremiumWord & "'," & vPremiumAmount & "," & vPremiumAfter & "," _
                            & "'" & vCommissionWord & "'," & vCommissionAmount & "," & vCommissionAfter & "," _
                            & "" & vGrossProfitPercent & "," & vGrossProfitAmount & "," & vInterestStockPercent & "," & vInterestStockAmount & "," _
                            & "" & vProfitPercent & "," & vProfitAmount & "," & vProfitPercent_W & "," & vProfitAmount_W & ",'" & vMyDescriptionSub & "'," & vFromQTY & "," & vToQTY & "," & vPrice1CashRec & "," & vPrice1CashDel & "," & vPrice1CreditRec & "," & vPrice1CreditDel & "," & vPriceSet2 & ", " _
                            & "" & vIsPriceUpdate & ",'" & vToUpdateDate & "'," & vIsUpdate & " "
                            vExecuteQuery = New SqlCommand(vQuery, vConnection)
                            vExecuteQuery.ExecuteNonQuery()
                            Me.PB101.Value = i
                        Next


                        vQuery = "commit tran"
                        vExecuteQuery = New SqlCommand(vQuery, vConnection)
                        vExecuteQuery.ExecuteNonQuery()

                        vQuery = "exec dbo.USP_PS_DeliverySendMail '" & vDocno & "'"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()

                        MsgBox("�ѹ�֡�������ç���ҧ�Ҥ��Ţ��� " & vDocno & " ���º�������Ǥ�Ѻ")
                        Me.PB101.Value = 0
                        Me.DGHeader.DataSource = Nothing
                        Me.DGDetails.DataSource = Nothing
                        Me.LBLFileName.Text = ""
                        Me.PB101.Value = 0
                        Me.TextDocNo.Text = ""
                        Me.TextMyDescription.Text = ""
                        Me.DocDate.Text = Now.Date
                        Me.TextDocNo.Text = ""
                        Me.PBNew.Visible = True
                        Me.PBConfirm.Visible = False
                        vCheckIsConfirm = 0

                        If Me.frmPriceStructureRequest Is Nothing Then
                            frmPriceStructureRequest = New FormPriceStructureRequest
                        Else
                            If frmPriceStructureRequest.IsDisposed Then
                                frmPriceStructureRequest = New FormPriceStructureRequest
                            End If
                        End If

                        vPriceStructureDocNo = Trim(vDocno)
                        frmPriceStructureRequest.Show()
                        frmPriceStructureRequest.BringToFront()


                    Catch ex As Exception
                        MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
                        vQuery = "rollback tran"
                        vExecuteQuery = New SqlCommand(vQuery, vConnection)
                        vExecuteQuery.ExecuteNonQuery()
                    End Try
                Else
                    MsgBox("�͡�������բ������������ö�ѹ�֡��������", MsgBoxStyle.Critical, "Send Message Information")
                End If
            Else
                MsgBox("�͡��÷��͹��ѵ������������ö��䢢�������", MsgBoxStyle.Critical, "Send Message Information")
            End If
        Else
            MsgBox("�͡����բ�ͼԴ��Ҵ�ͧ������ �������ö�ѹ�֡��������", MsgBoxStyle.Critical, "Send Message Information")
        End If

    End Sub

    Private Sub SaveData()
        'Dim vDocno As String
        'Dim vDocDate As String
        'Dim vTargetProfit As String
        'Dim vProfit As String
        'Dim vSmartPoint As String
        'Dim vMemberDiscount As String
        'Dim vFileDataSource As String
        'Dim vPathFile As String
        ''------------------------------------------
        'Dim vItemCode As String
        'Dim vItemName As String
        'Dim vSaleUnitCode As String
        'Dim vDO As Double
        'Dim vPriceSet As Double
        'Dim vDiscountBillWord As String
        'Dim vDiscountBill1 As Double
        'Dim vDiscountBillAmount As Double
        'Dim vAccCost As Double
        'Dim vDiscountFollow1Word As String
        'Dim vDiscountFollow11 As Double
        'Dim vDiscountFollow1Amount As Double
        'Dim vDiscountFollow1After As Double
        'Dim vDiscountFollow2Word As String
        'Dim vDiscountFollow21 As Double
        'Dim vDiscountFollow2Amount As Double
        'Dim vDiscountFollow2After As Double
        'Dim vDiscountFollow3Word As String
        'Dim vDiscountFollow31 As Double
        'Dim vDiscountFollow3Amount As Double
        'Dim vDiscountFollow3After As Double
        'Dim vDiscountFollow4Word As String
        'Dim vDiscountFollow41 As Double
        'Dim vDiscountFollow4Amount As Double
        'Dim vDiscountFollow4After As Double
        'Dim vDiscountRebateWord As String
        'Dim vDiscountRebate1 As Double
        'Dim vDiscountRebateAmount As Double
        'Dim vDiscountRebateAfter As Double
        'Dim vDiscountSpecialWord As String
        'Dim vDiscountSpecial1 As Double
        'Dim vDiscountSpecialAmount As Double
        'Dim vNetCost As Double
        'Dim vLossBudgetWord As String
        'Dim vLossBudget1 As Double
        'Dim vLossBudgetAmount As Double
        'Dim vLossBudgetAfter As Double
        'Dim vTransferInWord As String
        'Dim vTransferIn1 As Double
        'Dim vTransferOutWord As String
        'Dim vTransferOut1 As Double
        'Dim vAdvertiseWord As String
        'Dim vAdvertise1 As Double
        'Dim vAdvertiseAmount As Double
        'Dim vAdvertiseAfter As Double
        'Dim vVatCost As Double
        'Dim vVatAmount As Double
        'Dim vSetupWord As String
        'Dim vSetupAmount As Double
        'Dim vServiceWord As String
        'Dim vServiceAmount As Double
        'Dim vMarketCost As Double
        'Dim vPointWord As String
        'Dim vPoint1 As Double
        'Dim vPointAmount As Double
        'Dim vPointAfter As Double
        'Dim vTargetWord As String
        'Dim vTargetAmount As Double
        'Dim vPremiumWord As String
        'Dim vPremiumAmount As Double
        'Dim vCommissionWord As String
        'Dim vCommission1 As Double
        'Dim vCommissionAmount As Double
        'Dim vCommissionAfter As Double
        'Dim vGrossProfitPercent As String
        'Dim vGrossProfitAmount As Double
        'Dim vInterestStockPercent As String
        'Dim vInterestStockAmount As Double
        'Dim vProfitPercent As String
        'Dim vProfitAmount As Double
        'Dim vProfitPercent_W As String
        'Dim vProfitAmount_W As Double
        'Dim vMyDescription As String
        'Dim vMyDescriptionSub As String
        'Dim vTransferInAfter As Double
        'Dim vTransferOutAfter As Double
        'Dim vVatWord As String
        'Dim vSetupAfter As Double
        'Dim vTargetAfter As Double
        'Dim vPremiumAfter As Double
        '--------------------------------------------------------------

        '---------------------------------------------------------------
        'Dim i As Integer
        'Dim vLenFilePath As Integer
        'Dim vLenMark As Integer

        'Dim excel As New Microsoft.Office.Interop.Excel.Application
        'Dim wb As Microsoft.Office.Interop.Excel.Workbook

        'If vCheckErrorFileExcel = 0 Then
        '    If vCheckIsConfirm = 0 Then
        '        If Me.DGHeader.Rows.Count > 0 And Me.DGDetails.Rows.Count > 0 Then
        '            Me.PB101.Value = 1
        '            If Me.TextDocNo.Text = "" Then
        '                vQuery = "exec dbo.USP_PS_NewDocno"
        '                da = New SqlDataAdapter(vQuery, vConnection)
        '                ds = New DataSet
        '                da.Fill(ds, "NewDocno")
        '                vdt = ds.Tables("NewDocno")
        '                vDocno = vdt.Rows(0).Item("newdocno")
        '                vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        '            Else
        '                vDocno = Me.TextDocNo.Text
        '                vDocDate = Me.DocDate.Text
        '            End If


        '            Try
        '                vQuery = "begin tran"
        '                vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '                vExecuteQuery.ExecuteNonQuery()
        '                If Not IsDBNull(Me.DGHeader.Rows(0).Cells(1).Value) Then
        '                    vTargetProfit = CType((CType(Me.DGHeader.Rows(0).Cells(1).Value, Double) * 100), String) & "%"
        '                Else
        '                    vTargetProfit = 0
        '                End If
        '                If Not IsDBNull(Me.DGHeader.Rows(1).Cells(1).Value) Then
        '                    vProfit = CType((CType(Me.DGHeader.Rows(1).Cells(1).Value, Double) * 100), String) & "%"
        '                Else
        '                    vProfit = 0
        '                End If
        '                If Not IsDBNull(Me.DGHeader.Rows(2).Cells(1).Value) Then
        '                    vSmartPoint = CType((CType(Me.DGHeader.Rows(2).Cells(1).Value, Double) * 100), String) & "%"
        '                Else
        '                    vSmartPoint = 0
        '                End If
        '                If Not IsDBNull(Me.DGHeader.Rows(3).Cells(1).Value) Then
        '                    vMemberDiscount = CType((CType(Me.DGHeader.Rows(3).Cells(1).Value, Double) * 100), String) & "%"
        '                Else
        '                    vMemberDiscount = 0
        '                End If
        '                vMyDescription = Me.TextMyDescription.Text
        '                vPathFile = Trim(Me.LBLFileName.Text)
        '                vLenMark = InStr(Me.LBLFileName.Text, "@")
        '                vLenFilePath = Len(Me.LBLFileName.Text)
        '                vFileDataSource = Microsoft.VisualBasic.Right(Me.LBLFileName.Text, vLenFilePath - vLenMark)
        '                vQuery = "exec dbo.USP_PS_InsertPriceStructureSet1 '" & vDocno & "','" & vDocDate & "','" & vTargetProfit & "','" & vProfit & "','" & vSmartPoint & "','" & vMemberDiscount & "','" & vFileDataSource & "','" & vPathFile & "','" & vMyDescription & "'"
        '                vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '                vExecuteQuery.ExecuteNonQuery()

        '                Me.PB101.Maximum = Me.DGDetails.RowCount - 2

        '                For i = 0 To Me.DGDetails.RowCount - 2

        '                    vItemCode = Trim(Me.DGDetails.Rows(i).Cells(0).Value)
        '                    vItemName = Trim(Me.DGDetails.Rows(i).Cells(1).Value)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(2).Value) Then
        '                        vSaleUnitCode = Trim(Me.DGDetails.Rows(i).Cells(2).Value)
        '                    Else
        '                        vSaleUnitCode = ""
        '                    End If
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(3).Value) Then
        '                        vDO = Trim(Me.DGDetails.Rows(i).Cells(3).Value)
        '                    Else
        '                        vDO = 0
        '                    End If
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(42).Value) Then
        '                        vPriceSet = Me.DGDetails.Rows(i).Cells(42).Value
        '                    Else
        '                        vPriceSet = 0
        '                    End If
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(4).Value) Then
        '                        vDiscountBillWord = CType((CType(Me.DGDetails.Rows(i).Cells(4).Value, Double) * 100), String) & "%"
        '                        vDiscountBill1 = (CType(Me.DGDetails.Rows(i).Cells(4).Value, Double) * 100)
        '                    Else
        '                        vDiscountBillWord = ""
        '                        vDiscountBill1 = 0
        '                    End If
        '                    vDiscountBillAmount = CalcAmount(vDO, vDiscountBill1)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(5).Value) Then
        '                        vAccCost = Me.DGDetails.Rows(i).Cells(5).Value
        '                    Else
        '                        vAccCost = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(6).Value) Then
        '                        vDiscountFollow1Word = CType((CType(Me.DGDetails.Rows(i).Cells(6).Value, Double) * 100), String) & "%"
        '                        vDiscountFollow11 = (CType(Me.DGDetails.Rows(i).Cells(6).Value, Double) * 100)
        '                    Else
        '                        vDiscountFollow1Word = ""
        '                        vDiscountFollow11 = 0
        '                    End If
        '                    vDiscountFollow1Amount = CalcAmount(vAccCost, vDiscountFollow11)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(7).Value) Then
        '                        vDiscountFollow1After = Me.DGDetails.Rows(i).Cells(7).Value
        '                    Else
        '                        vDiscountFollow1After = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(8).Value) Then
        '                        vDiscountFollow2Word = CType((CType(Me.DGDetails.Rows(i).Cells(8).Value, Double) * 100), String) & "%"
        '                        vDiscountFollow21 = (CType(Me.DGDetails.Rows(i).Cells(8).Value, Double) * 100)
        '                    Else
        '                        vDiscountFollow2Word = ""
        '                        vDiscountFollow21 = 0
        '                    End If
        '                    vDiscountFollow2Amount = CalcAmount(vDiscountFollow1After, vDiscountFollow21)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(9).Value) Then
        '                        vDiscountFollow2After = Me.DGDetails.Rows(i).Cells(9).Value
        '                    Else
        '                        vDiscountFollow2After = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(10).Value) Then
        '                        vDiscountFollow3Word = CType((CType(Me.DGDetails.Rows(i).Cells(10).Value, Double) * 100), String) & "%"
        '                        vDiscountFollow31 = (CType(Me.DGDetails.Rows(i).Cells(10).Value, Double) * 100)
        '                    Else
        '                        vDiscountFollow3Word = ""
        '                        vDiscountFollow31 = 0
        '                    End If
        '                    vDiscountFollow3Amount = CalcAmount(vDiscountFollow2After, vDiscountFollow31)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(11).Value) Then
        '                        vDiscountFollow3After = Me.DGDetails.Rows(i).Cells(11).Value
        '                    Else
        '                        vDiscountFollow3After = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(12).Value) Then
        '                        vDiscountFollow4Word = CType((CType(Me.DGDetails.Rows(i).Cells(12).Value, Double) * 100), String) & "%"
        '                        vDiscountFollow41 = (CType(Me.DGDetails.Rows(i).Cells(12).Value, Double) * 100)
        '                    Else
        '                        vDiscountFollow4Word = ""
        '                        vDiscountFollow41 = 0
        '                    End If
        '                    vDiscountFollow4Amount = CalcAmount(vDiscountFollow3After, vDiscountFollow41)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(13).Value) Then
        '                        vDiscountFollow4After = Me.DGDetails.Rows(i).Cells(13).Value
        '                    Else
        '                        vDiscountFollow4After = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(14).Value) Then
        '                        vDiscountRebateWord = CType((CType(Me.DGDetails.Rows(i).Cells(14).Value, Double) * 100), String) & "%"
        '                        vDiscountRebate1 = (CType(Me.DGDetails.Rows(i).Cells(14).Value, Double) * 100)
        '                    Else
        '                        vDiscountRebateWord = ""
        '                        vDiscountRebate1 = 0
        '                    End If
        '                    vDiscountRebateAmount = CalcAmount(vDiscountFollow4After, vDiscountRebate1)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(15).Value) Then
        '                        vDiscountRebateAfter = Me.DGDetails.Rows(i).Cells(15).Value
        '                    Else
        '                        vDiscountRebateAfter = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(16).Value) Then
        '                        vDiscountSpecialWord = Me.DGDetails.Rows(i).Cells(16).Value
        '                        vDiscountSpecial1 = Me.DGDetails.Rows(i).Cells(16).Value
        '                    Else
        '                        vDiscountSpecialWord = 0
        '                        vDiscountSpecial1 = 0
        '                    End If
        '                    vDiscountSpecialAmount = vDiscountSpecial1
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(17).Value) Then
        '                        vNetCost = Me.DGDetails.Rows(i).Cells(17).Value
        '                    Else
        '                        vNetCost = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(18).Value) Then
        '                        vLossBudgetWord = CType((CType(Me.DGDetails.Rows(i).Cells(18).Value, Double) * 100), String) & "%"
        '                        vLossBudget1 = (CType(Me.DGDetails.Rows(i).Cells(18).Value, Double) * 100)
        '                    Else
        '                        vLossBudgetWord = ""
        '                        vLossBudget1 = 0
        '                    End If
        '                    vLossBudgetAmount = CalcAmount(vNetCost, vLossBudget1)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(19).Value) Then
        '                        vLossBudgetAfter = Me.DGDetails.Rows(i).Cells(19).Value
        '                    Else
        '                        vLossBudgetAfter = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(20).Value) Then
        '                        vTransferInWord = Me.DGDetails.Rows(i).Cells(20).Value
        '                        vTransferIn1 = Me.DGDetails.Rows(i).Cells(20).Value
        '                        vTransferInAfter = vLossBudgetAfter + vTransferIn1
        '                    Else
        '                        vTransferIn1 = 0
        '                        vTransferInWord = ""
        '                        vTransferInAfter = vLossBudgetAfter + vTransferIn1
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(21).Value) Then
        '                        vTransferOutWord = Me.DGDetails.Rows(i).Cells(21).Value
        '                        vTransferOut1 = Me.DGDetails.Rows(i).Cells(21).Value
        '                        vTransferOutAfter = vTransferInAfter + vTransferOut1
        '                    Else
        '                        vTransferOutWord = ""
        '                        vTransferOut1 = 0
        '                        vTransferOutAfter = vTransferInAfter + vTransferOut1
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(22).Value) Then
        '                        vAdvertiseWord = CType((CType(Me.DGDetails.Rows(i).Cells(22).Value, Double) * 100), String) & "%"
        '                        vAdvertise1 = (CType(Me.DGDetails.Rows(i).Cells(22).Value, Double) * 100)
        '                    Else
        '                        vAdvertiseWord = ""
        '                        vAdvertise1 = 0
        '                    End If
        '                    vAdvertiseAmount = CalcAmount(vTransferOutAfter, vAdvertise1)
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(23).Value) Then
        '                        vAdvertiseAfter = Me.DGDetails.Rows(i).Cells(23).Value
        '                    Else
        '                        vAdvertiseAfter = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(24).Value) Then
        '                        vVatCost = Me.DGDetails.Rows(i).Cells(24).Value
        '                        vVatAmount = (vAdvertiseAfter * 7) / 100
        '                        vVatWord = "7%"
        '                    Else
        '                        vVatCost = 0
        '                        vVatAmount = 0
        '                        vVatWord = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(25).Value) Then
        '                        vSetupWord = Me.DGDetails.Rows(i).Cells(25).Value
        '                        vSetupAmount = Me.DGDetails.Rows(i).Cells(25).Value
        '                        vSetupAfter = vVatCost + vSetupAmount
        '                    Else
        '                        vSetupWord = ""
        '                        vSetupAmount = 0
        '                        vSetupAfter = vVatCost + vSetupAmount
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(26).Value) Then
        '                        vServiceWord = Me.DGDetails.Rows(i).Cells(26).Value
        '                        vServiceAmount = Me.DGDetails.Rows(i).Cells(26).Value
        '                    Else
        '                        vServiceWord = ""
        '                        vServiceAmount = 0
        '                    End If
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(27).Value) Then
        '                        vMarketCost = Me.DGDetails.Rows(i).Cells(27).Value
        '                    Else
        '                        vMarketCost = 0
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(34).Value) Then
        '                        vPointWord = CType((CType(Me.DGDetails.Rows(i).Cells(34).Value, Double) * 100), String) & "%"
        '                    Else
        '                        vPointWord = ""
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(35).Value) Then
        '                        vPoint1 = Me.DGDetails.Rows(i).Cells(35).Value
        '                    Else
        '                        vPoint1 = 0
        '                    End If
        '                    vPointAmount = vPoint1
        '                    vPointAfter = vMarketCost + vPointAmount

        '                    If vPointWord = "" And vPointAmount = 0 Then
        '                        MsgBox("�����Թ��� " & vItemCode & "   " & vItemName & " ����դ�Ңͧ Smart Point �������ö�ѹ�֡��������  ��س���䢢����š�͹�ѹ�֡����")
        '                        vQuery = "rollback tran"
        '                        vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '                        vExecuteQuery.ExecuteNonQuery()
        '                        Exit Sub
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(36).Value) Then
        '                        vTargetWord = Me.DGDetails.Rows(i).Cells(36).Value
        '                        vTargetAmount = Me.DGDetails.Rows(i).Cells(36).Value
        '                        vTargetAfter = vPointAfter + vTargetAmount
        '                    Else
        '                        vTargetWord = ""
        '                        vTargetAmount = 0
        '                        vTargetAfter = vPointAfter + vTargetAmount
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(37).Value) Then
        '                        vPremiumWord = Me.DGDetails.Rows(i).Cells(37).Value
        '                        vPremiumAmount = Me.DGDetails.Rows(i).Cells(37).Value
        '                        vPremiumAfter = vTargetAfter + vPremiumAmount
        '                    Else
        '                        vPremiumWord = ""
        '                        vPremiumAmount = 0
        '                        vPremiumAfter = vTargetAfter + vPremiumAmount
        '                    End If

        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(38).Value) Then
        '                        vCommissionWord = CType((CType(Me.DGDetails.Rows(i).Cells(38).Value, Double) * 100), String) & "%"
        '                    Else
        '                        vCommissionWord = ""
        '                    End If
        '                    vCommissionAmount = vCommission1
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(39).Value) Then
        '                        vCommission1 = Me.DGDetails.Rows(i).Cells(39).Value
        '                    Else
        '                        vCommission1 = 0
        '                    End If
        '                    vCommissionAmount = vCommission1
        '                    vCommissionAfter = vPremiumAfter + vCommissionAmount

        '                    vGrossProfitPercent = (Me.DGDetails.Rows(i).Cells(48).Value * 100)
        '                    vGrossProfitAmount = ((Me.DGDetails.Rows(i).Cells(48).Value * 100) * vMarketCost) / 100
        '                    vInterestStockPercent = (Me.DGDetails.Rows(i).Cells(28).Value * 100)
        '                    vInterestStockAmount = Me.DGDetails.Rows(i).Cells(29).Value
        '                    vProfitPercent = (Me.DGDetails.Rows(i).Cells(30).Value * 100)
        '                    vProfitAmount = Me.DGDetails.Rows(i).Cells(31).Value
        '                    vProfitPercent_W = (Me.DGDetails.Rows(i).Cells(32).Value * 100)
        '                    vProfitAmount_W = Me.DGDetails.Rows(i).Cells(33).Value
        '                    If Not IsDBNull(Me.DGDetails.Rows(i).Cells(47).Value) Then
        '                        vMyDescriptionSub = Me.DGDetails.Rows(i).Cells(47).Value
        '                    Else
        '                        vMyDescriptionSub = ""
        '                    End If
        '                    '---------------------------------------------------------------------------------------------------
        '                    Dim vFromQTY As Double
        '                    Dim vToQTY As Double
        '                    Dim vPriceSet2 As Double
        '                    Dim vIsPriceUpdate As Integer = 1
        '                    Dim vToUpdateDate As String = Me.DGDetails.Rows(i).Cells(46).Value
        '                    Dim vIsUpdate As Integer = 0
        '                    Dim vIsPrintLabel As Integer = 0
        '                    Dim vPrice1CashRec As Double
        '                    Dim vPrice1CashDel As Double
        '                    Dim vPrice1CreditRec As Double
        '                    Dim vPrice1CreditDel As Double
        '                    '---------------------------------------------------------------------------------------------------

        '                    vSaleUnitCode = Me.DGDetails.Rows(i).Cells(2).Value
        '                    vFromQTY = 1
        '                    vToQTY = 99999
        '                    vPrice1CashRec = Me.DGDetails.Rows(i).Cells(41).Value
        '                    vPrice1CashDel = Me.DGDetails.Rows(i).Cells(42).Value
        '                    vPrice1CreditRec = Me.DGDetails.Rows(i).Cells(43).Value
        '                    vPrice1CreditDel = Me.DGDetails.Rows(i).Cells(44).Value
        '                    vPriceSet2 = Me.DGDetails.Rows(i).Cells(45).Value


        '                    vQuery = "exec dbo.USP_PS_InsertPriceStructureSubSet '" & vDocno & "','" & vItemCode & "','" & vItemName & "','" & vSaleUnitCode & "'," & vDO & ", " _
        '                    & "" & vPriceSet & ",'" & vDiscountBillWord & "'," & vDiscountBillAmount & "," & vAccCost & "," _
        '                    & "'" & vDiscountFollow1Word & "'," & vDiscountFollow1Amount & "," & vDiscountFollow1After & "," _
        '                    & "'" & vDiscountFollow2Word & "'," & vDiscountFollow2Amount & "," & vDiscountFollow2After & "," _
        '                    & "'" & vDiscountFollow3Word & "'," & vDiscountFollow3Amount & "," & vDiscountFollow3After & "," _
        '                    & "'" & vDiscountFollow4Word & "'," & vDiscountFollow4Amount & "," & vDiscountFollow4After & "," _
        '                    & "'" & vDiscountRebateWord & "'," & vDiscountRebateAmount & "," & vDiscountRebateAfter & "," _
        '                    & "'" & vDiscountSpecialWord & "'," & vDiscountSpecialAmount & "," & vNetCost & "," _
        '                    & "'" & vLossBudgetWord & "'," & vLossBudgetAmount & "," & vLossBudgetAfter & "," _
        '                    & "'" & vTransferInWord & "'," & vTransferIn1 & "," & vTransferInAfter & "," _
        '                    & "'" & vTransferOutWord & "'," & vTransferOut1 & "," & vTransferOutAfter & "," _
        '                    & "'" & vAdvertiseWord & "'," & vAdvertiseAmount & "," & vAdvertiseAfter & "," _
        '                    & "'" & vVatWord & "'," & vVatCost & "," & vVatAmount & "," _
        '                    & "'" & vSetupWord & "'," & vSetupAmount & "," & vSetupAfter & "," _
        '                    & "'" & vServiceWord & "'," & vServiceAmount & "," & vMarketCost & "," _
        '                    & "'" & vPointWord & "'," & vPointAmount & "," & vPointAfter & "," _
        '                    & "'" & vTargetWord & "'," & vTargetAmount & "," & vTargetAfter & "," _
        '                    & "'" & vPremiumWord & "'," & vPremiumAmount & "," & vPremiumAfter & "," _
        '                    & "'" & vCommissionWord & "'," & vCommissionAmount & "," & vCommissionAfter & "," _
        '                    & "" & vGrossProfitPercent & "," & vGrossProfitAmount & "," & vInterestStockPercent & "," & vInterestStockAmount & "," _
        '                    & "" & vProfitPercent & "," & vProfitAmount & "," & vProfitPercent_W & "," & vProfitAmount_W & ",'" & vMyDescriptionSub & "'," & vFromQTY & "," & vToQTY & "," & vPrice1CashRec & "," & vPrice1CashDel & "," & vPrice1CreditRec & "," & vPrice1CreditDel & "," & vPriceSet2 & ", " _
        '                    & "" & vIsPriceUpdate & ",'" & vToUpdateDate & "'," & vIsUpdate & " "
        '                    vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '                    vExecuteQuery.ExecuteNonQuery()
        '                    Me.PB101.Value = i
        '                Next


        '                vQuery = "commit tran"
        '                vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '                vExecuteQuery.ExecuteNonQuery()

        '                vQuery = "exec dbo.USP_PS_DeliverySendMail '" & vDocno & "'"
        '                vCMD = New SqlCommand(vQuery, vConnection)
        '                vCMD.ExecuteNonQuery()

        '                If Me.TextDocNo.Text = "" Then
        '                    MsgBox("�ѹ�֡�������ç���ҧ�Ҥ��Ţ��� " & vDocno & " ���º�������Ǥ�Ѻ")
        '                    Me.PB101.Value = 0
        '                    Me.DGHeader.DataSource = Nothing
        '                    Me.DGDetails.DataSource = Nothing
        '                    Me.LBLFileName.Text = ""
        '                    Me.PB101.Value = 0
        '                    Me.TextDocNo.Text = ""
        '                    Me.TextMyDescription.Text = ""
        '                    Me.DocDate.Text = Now.Date
        '                    Me.TextDocNo.Text = ""
        '                    Me.PBNew.Visible = True
        '                    Me.PBConfirm.Visible = False
        '                    vCheckIsConfirm = 0
        '                End If

        '                If Me.frmPriceStructureRequest Is Nothing Then
        '                    frmPriceStructureRequest = New FormPriceStructureRequest
        '                Else
        '                    If frmPriceStructureRequest.IsDisposed Then
        '                        frmPriceStructureRequest = New FormPriceStructureRequest
        '                    End If
        '                End If

        '                Dim vAnswerPrint As Integer

        '                vAnswerPrint = MsgBox("��ͧ��þ������ʹ��ç���ҧ�Ҥ��������", MsgBoxStyle.YesNo, "Send Question Message")

        '                If vAnswerPrint = 6 Then

        '                    Me.PB101.Value = 0
        '                    Me.DGHeader.DataSource = Nothing
        '                    Me.DGDetails.DataSource = Nothing
        '                    Me.LBLFileName.Text = ""
        '                    Me.PB101.Value = 0
        '                    Me.TextDocNo.Text = ""
        '                    Me.TextMyDescription.Text = ""
        '                    Me.DocDate.Text = Now.Date
        '                    Me.TextDocNo.Text = ""
        '                    Me.PBNew.Visible = True
        '                    Me.PBConfirm.Visible = False
        '                    vCheckIsConfirm = 0

        '                    vPriceStructureDocNo = Trim(vDocno)
        '                    frmPriceStructureRequest.Show()
        '                    frmPriceStructureRequest.BringToFront()
        '                End If


        '            Catch ex As Exception
        '                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '                vQuery = "rollback tran"
        '                vExecuteQuery = New SqlCommand(vQuery, vConnection)
        '                vExecuteQuery.ExecuteNonQuery()
        '            End Try
        '        Else
        '            MsgBox("�͡�������բ������������ö�ѹ�֡��������", MsgBoxStyle.Critical, "Send Message Information")
        '        End If
        '    Else
        '        MsgBox("�͡��÷��͹��ѵ������������ö��䢢�������", MsgBoxStyle.Critical, "Send Message Information")
        '    End If
        'Else
        '    MsgBox("�͡����բ�ͼԴ��Ҵ�ͧ������ �������ö�ѹ�֡��������", MsgBoxStyle.Critical, "Send Message Information")
        'End If

    End Sub

    Private Sub BTNGenData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGenData.Click
        'On Error GoTo ErrDescription

        If Me.LBLFileName.Text <> "" Then
            If vPathFileName <> "" Then
                If InStr(vPathFileName, "@") > 0 Then
                    Call GetDataHeaderFromExcel()
                    Call GetDataDetailsFromExcel()
                    MsgBox("��ҵ�ͧ�������¹�����ŷ�������㹡�û�Ѻ�Ҥ��Թ������� Excel ��سҡ������ѹ�֡�ա����", MsgBoxStyle.Information, "Send Information")
                    If Me.DGDetails.RowCount > 0 And Me.TextDocNo.Text <> "" Then
                        Call SaveData()
                    End If
                Else
                    MsgBox("����ç���ҧ�Ҥҵ�ͧ��駪�������� ����ͧ���� @ ��ҧ˹�Ҫ���������  ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            Else
                MsgBox("���������ç���ҧ�Ҥҷ��д֧������  ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If


        'ErrDescription:
        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If

    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Me.DGHeader.DataSource = ""
        Me.DGDetails.DataSource = ""
        Me.PB101.Value = 0

    End Sub

    Public Function CalcAmount(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmount = (vPriceSetAmount * vPercent) / 100
    End Function

    Public Function CalcAmountAfterAdd(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmountAfterAdd = vPriceSetAmount + (vPriceSetAmount * vPercent) / 100
    End Function
    Public Function CalcAmountAfterDelete(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmountAfterDelete = vPriceSetAmount - (vPriceSetAmount * vPercent) / 100
    End Function


Private Sub BTNFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNFind.Click
        Dim vCheckDocno As String
        Dim vLenMark As Integer
        Dim vLenFilePath As Integer
        Dim vFileName As String
        Dim vPath As String

        On Error Resume Next

        Me.OpenFileDialog1.InitialDirectory = "Q:\RP\�Ѵ����\�ç���ҧ�Ҥ�\"
        OpenFileDialog1.Filter = "Xls Files (*.xls)|*.xls" & "Xls Files(*.xlsx)|*.xlsx"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            LBLFileName.Text = OpenFileDialog1.FileName
            vPathFileName = Me.LBLFileName.Text
            vLenMark = InStr(Me.LBLFileName.Text, "@")
            vLenFilePath = Len(Me.LBLFileName.Text)
            vPath = Microsoft.VisualBasic.Left(Me.LBLFileName.Text, vLenMark - 1)
            vFileName = Microsoft.VisualBasic.Right(Me.LBLFileName.Text, vLenFilePath - vLenMark)


            vQuery = "select top 1 isnull(docno,'') as docno,docdate ,isnull(mydescription,'') as mydescription,isnull(isconfirm,0) as isconfirm  from npmaster.dbo.TB_PS_PriceStructure where datasource = '" & vFileName & "'and iscancel = 0 "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "CheckDocNo")
            dt = ds.Tables("CheckDocNo")
            If dt.Rows.Count > 0 Then
                vCheckDocno = dt.Rows(0).Item("docno")
                Me.TextDocNo.Text = vCheckDocno
                Me.TextMyDescription.Text = dt.Rows(0).Item("mydescription")
                Me.DocDate.Value = dt.Rows(0).Item("docdate")
                vCheckIsConfirm = dt.Rows(0).Item("isconfirm")
                If vCheckIsConfirm = 0 Then
                    Me.PBNew.Visible = True
                    Me.PBConfirm.Visible = False
                Else
                    Me.PBNew.Visible = False
                    Me.PBConfirm.Visible = True
                End If
            Else
                Me.TextDocNo.Text = ""
                Me.TextMyDescription.Text = ""
                Me.DocDate.Value = Date.Now
                Me.PBNew.Visible = True
                Me.PBConfirm.Visible = False
                vCheckIsConfirm = 0
            End If
        End If
    End Sub

    Private Sub TextPriceStructureSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextPriceStructureSearch.KeyDown
        Dim vSearch As String
        Dim vListDocno As ListViewItem
        Dim i As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.TextPriceStructureSearch.Text <> "" Then
                vSearch = Me.TextPriceStructureSearch.Text
                vQuery = "exec dbo.USP_PS_SearchPriceStructure '" & vSearch & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "SearchDocno")
                dt = ds.Tables("SearchDocno")

                If dt.Rows.Count > 0 Then
                    Me.ListViewPriceStructure.Items.Clear()
                    For i = 0 To dt.Rows.Count - 1
                        vListDocno = Me.ListViewPriceStructure.Items.Add(dt.Rows(i).Item("docno"))
                        vListDocno.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                        vListDocno.SubItems.Add(2).Text = dt.Rows(i).Item("creatorcode")
                        vListDocno.SubItems.Add(3).Text = dt.Rows(i).Item("datasource")
                        vListDocno.SubItems.Add(4).Text = dt.Rows(i).Item("isconfirm")

                        If ListViewPriceStructure.Items(i).SubItems(4).Text = 0 Then
                            ListViewPriceStructure.Items(i).BackColor = Color.White
                        Else
                            ListViewPriceStructure.Items(i).BackColor = Color.Green
                            ListViewPriceStructure.Items(i).Checked = False
                        End If
                    Next
                    Me.ListViewPriceStructure.Items(0).Selected = True
                    Me.ListViewPriceStructure.HideSelection = False
                    Me.ListViewPriceStructure.MultiSelect = False
                    Me.ListViewPriceStructure.Focus()
                Else
                    Me.ListViewPriceStructure.Items.Clear()
                    Me.TextPriceStructureSearch.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub TextPriceStructureSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextPriceStructureSearch.TextChanged

    End Sub

    Private Sub BTNSearchDocno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchDocno.Click
        Me.GBSearchPriceStructure.Visible = True
        Me.TextPriceStructureSearch.Focus()
    End Sub

    Private Sub BTNPriceStructureConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPriceStructureConfirm.Click
        Dim i As Integer
        Dim vDocNo As String

        On Error Resume Next

        If Me.ListViewPriceStructure.Items.Count > 0 Then
            For i = 0 To Me.ListViewPriceStructure.Items.Count - 1
                If Me.ListViewPriceStructure.Items(i).Checked = True And Me.ListViewPriceStructure.Items(i).SubItems(4).Text <> 1 Then
                    vDocNo = Trim(Me.ListViewPriceStructure.Items(i).Text)

                    vQuery = "exec dbo.USP_PS_ConfirmForExcel '" & vDocNo & "'"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_PS_UpdateMasterForExcel '" & vDocNo & "'"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_PS_CKUpdateMasterForExcel '" & vDocNo & "'"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_PS_DeliverySendMail '" & vDocNo & "'"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    'Call MoveExcelConfirm(i)
                End If
            Next
            Me.ListViewPriceStructure.Items.Clear()
            MsgBox("͹��ѵ��͡��õ��������͡������º��������", MsgBoxStyle.Information, "Send Message Information")
            Me.TextPriceStructureSearch.Text = ""
        End If
    End Sub

    Private Sub MoveExcelConfirm(ByVal i As Integer)
        Dim vPathFileName As String
        Dim vPath As String
        Dim fileName As String
        Dim vLenPath As Integer
        Dim vRemark As Integer
        Dim vExcelRemark As Integer
        Dim vExcelFile As String

        If Me.ListViewPriceStructure.Items.Count > 0 Then
            vPathFileName = Me.ListViewPriceStructure.Items(i).SubItems(5).Text
            vLenPath = Len(vPathFileName)
            vRemark = InStr(vPathFileName, "@")
            vExcelRemark = InStr(vPathFileName, ".xls")
            vPath = Microsoft.VisualBasic.Left(vPathFileName, vRemark - 1)
            vExcelFile = Microsoft.VisualBasic.Left(vPathFileName, vExcelRemark - 1)
            Dim dir As New DirectoryInfo(vPath)
            For Each f As FileInfo In dir.GetFiles
                If (f.Extension = ".xls") Then
                    If vPathFileName = f.FullName Then
                        'fileName = f.FullName
                        fileName = vPathFileName
                        f.MoveTo(fileName.Substring(0, fileName.Length - 4))
                        fileName = Path.GetFileName(fileName)
                        Dim Newpaht As String = "Q:\RP\�Ѵ����\�ç���ҧ�Ҥ�\�ç���ҧ�Ҥ�͹��ѵ�����\"
                        Try
                            f.MoveTo(Newpaht & fileName)
                            Exit Sub
                        Catch
                            MsgBox("���������")
                        End Try
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub BTNPriceStructureExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPriceStructureExit.Click
        'Me.GBSearchPriceStructure.Visible = False
        Me.Close()
    End Sub

    Private Sub BTNPriceStructureSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPriceStructureSearch.Click
        Dim vSearch As String
        Dim vListDocno As ListViewItem
        Dim i As Integer

        On Error Resume Next

        If Me.TextPriceStructureSearch.Text <> "" Then
            vSearch = Me.TextPriceStructureSearch.Text
            vQuery = "exec dbo.USP_PS_SearchPriceStructure '" & vSearch & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "SearchDocno")
            dt = ds.Tables("SearchDocno")

            If dt.Rows.Count > 0 Then
                Me.ListViewPriceStructure.Items.Clear()
                For i = 0 To dt.Rows.Count - 1
                    vListDocno = Me.ListViewPriceStructure.Items.Add(dt.Rows(i).Item("docno"))
                    vListDocno.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                    vListDocno.SubItems.Add(2).Text = dt.Rows(i).Item("creatorcode")
                    vListDocno.SubItems.Add(3).Text = dt.Rows(i).Item("mydescription")
                    vListDocno.SubItems.Add(4).Text = dt.Rows(i).Item("isconfirm")
                    vListDocno.SubItems.Add(5).Text = dt.Rows(i).Item("datasource")

                    If ListViewPriceStructure.Items(i).SubItems(4).Text = 0 Then
                        ListViewPriceStructure.Items(i).BackColor = Color.White
                    Else
                        ListViewPriceStructure.Items(i).BackColor = Color.Green
                        ListViewPriceStructure.Items(i).Checked = False
                    End If
                Next
                Me.ListViewPriceStructure.Items(0).Selected = True
                Me.ListViewPriceStructure.HideSelection = False
                Me.ListViewPriceStructure.MultiSelect = False
                Me.ListViewPriceStructure.Focus()
            Else
                Me.ListViewPriceStructure.Items.Clear()
                Me.TextPriceStructureSearch.Focus()
            End If
        End If
    End Sub

    Private Sub BTNClearData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearData.Click
        Me.PB101.Value = 0
        Me.DGHeader.DataSource = Nothing
        Me.DGDetails.DataSource = Nothing
        Me.LBLFileName.Text = ""
        Me.PB101.Value = 0
        Me.TextDocNo.Text = ""
        Me.TextMyDescription.Text = ""
        Me.DocDate.Text = Now.Date
        Me.PBNew.Visible = True
        Me.PBConfirm.Visible = False
        vCheckIsConfirm = 0
        vPathFileName = ""
        vCheckItemExist = 0
        Me.TextDocNo.Text = ""
    End Sub

    Private Sub BTNPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        On Error Resume Next

        If Me.TextDocNo.Text <> "" Then
            If Me.DGDetails.Rows.Count > 0 And Me.DGHeader.Rows.Count > 0 Then

                If Me.frmPriceStructureRequest Is Nothing Then
                    frmPriceStructureRequest = New FormPriceStructureRequest
                Else
                    If frmPriceStructureRequest.IsDisposed Then
                        frmPriceStructureRequest = New FormPriceStructureRequest
                    End If
                End If

                vPriceStructureDocNo = Trim(Me.TextDocNo.Text)
                frmPriceStructureRequest.Show()
                frmPriceStructureRequest.BringToFront()
            Else
                MsgBox("�������ö������͡����� ���ͧ�ҡ�ѧ����� Generate ������", MsgBoxStyle.Critical, "Send Error")
            End If
        Else
            MsgBox("�������ö������͡����� ���ͧ�ҡ�ѧ����ա�úѹ�֡������", MsgBoxStyle.Critical, "Send Error")
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        'Call MoveExcelConfirm(1)
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vDO As Double
        Dim vDiscount As Integer
        'Dim vAmount As Double
        Dim vCalcString As String

        vDO = 2000
        vDiscount = 3

        vQuery = "select acccost from npmaster.dbo.TB_NP_PriceStructureCalc"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchDocno")
        dt = ds.Tables("SearchDocno")

        If dt.Rows.Count > 0 Then
            vCalcString = dt.Rows(0).Item("acccost")
        End If

        'vAmount = vCalcString

    End Sub
End Class

