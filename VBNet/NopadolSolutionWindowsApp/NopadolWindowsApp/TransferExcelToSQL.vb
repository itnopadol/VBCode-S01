Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class TransferExcelToSQL
    Dim dt As DataTable
    Dim vList As ListViewItem
    Dim i As Integer
    Dim vFileName As String

    Private Sub TransferExcelToSQL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim DtSet As System.Data.DataSet

        'Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        'Dim MyConnection As System.Data.OleDb.OleDbConnection

        'Try
        '    Opn.Filter = "Excel Files (*.xls)|*.xls"     'เปิด excel
        '    Opn.ShowDialog()
        '    lblFilePath.Text = Opn.FileName
        'Catch ex As Exception

        'End Try


        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; data source='D:\Documents\Excel\โครงสร้างราคา_ประตูหน้าต่างและอุปกรณ์ 07022007.xls'; " & "Extended Properties=Excel 8.0;")

        '' เลือกข้อมูลจาก Sheet1 ครับ

        ''MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [sheet1$]", MyConnection)
        'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Cement$]", MyConnection)



        'MyCommand.TableMappings.Add("Table", "Attendence")

        'DtSet = New System.Data.DataSet

        'MyCommand.Fill(DtSet, "Data")
        'dt = DtSet.Tables("Data")
        'If dt.Rows.Count > 0 Then
        '    For i = 0 To dt.Rows.Count - 1
        '        vList = Me.ListView1.Items.Add(dt.Rows(i).Item("ยี่ห้อ"))
        '        vList.SubItems.Add(1).Text = dt.Rows(i).Item("รหัสสินค้า")
        '        'vList.SubItems.Add(2).Text = dt.Rows(i).Item("คลัง")
        '        'vList.SubItems.Add(3).Text = dt.Rows(i).Item("ชั้นเก็บ")
        '        'vList.SubItems.Add(4).Text = dt.Rows(i).Item("คงเหลือ")
        '        'vList.SubItems.Add(5).Text = dt.Rows(i).Item("หน่วย")
        '    Next
        'End If
        'MyConnection.Close()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.OpenFileDialog1.InitialDirectory = "Q:\RP\จัดซื้อ\โครงสร้างราคา\"
        OpenFileDialog1.Filter = "Excel Files (*.xls)|*.xls"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Label2.Text = OpenFileDialog1.FileName
            vFileName = Trim(Me.Label2.Text)
            'Call GenData()
            Call SetExcel()
        End If
    End Sub

    Public Sub GenData()
        Dim DtSet As System.Data.DataSet

        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection


        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; data source='" & vFileName & "'; " & "Extended Properties=Excel 8.0;")

        ' เลือกข้อมูลจาก Sheet1 ครับ

        'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [sheet1$]", MyConnection)
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [ลูกบิด$]", MyConnection)



        MyCommand.TableMappings.Add("Table", "Attendence")

        DtSet = New System.Data.DataSet

        MyCommand.Fill(DtSet, "Data")
        dt = DtSet.Tables("Data")
        DataGridView1.DataSource = dt
        'DataGridView1.DataBindings()
        'If dt.Rows.Count > 0 Then
        '    For i = 0 To dt.Rows.Count - 1
        '        vList = Me.ListView1.Items.Add(dt.Rows(i).Item("ยี่ห้อ"))
        '        vList.SubItems.Add(1).Text = dt.Rows(i).Item("รหัสสินค้า")
        '        'vList.SubItems.Add(2).Text = dt.Rows(i).Item("คลัง")
        '        'vList.SubItems.Add(3).Text = dt.Rows(i).Item("ชั้นเก็บ")
        '        'vList.SubItems.Add(4).Text = dt.Rows(i).Item("คงเหลือ")
        '        'vList.SubItems.Add(5).Text = dt.Rows(i).Item("หน่วย")
        '    Next
        'End If
        MyConnection.Close()
    End Sub

    Private Sub SetExcel()
        Dim ex As New excel.Application
        Dim wb As excel.Workbook
        Dim ws As excel.Worksheet
        ' ---- เปิด excel worksheet.        
        wb = ex.Workbooks.Open(vFileName)
        ex.ActiveWorkbook.Sheets(1).Select()
        ws = ex.ActiveWorkbook.ActiveSheet
        ' ---- สร้าง DataTable        
        Dim row As Integer
        Dim HasData As Boolean = True
        Dim dt As New DataTable("SheetOne")
        Dim dr As DataRow
        dt.Columns.Add("รหัสสินค้า", GetType(String))
        dt.Columns.Add("y", GetType(String))
        dt.Columns.Add("z", GetType(String))

        row = 9 ' ---- ข้อมูลเริ่มที่แถวที่ 2 (แถวแรกคือ 1, คอลัมน์แรกคือ 1)        
        Do
            If (IsNumeric(ws.Cells(row, 1).value)) Then                ' ---- ถ้าข้อมูลในคอลัมน์แรกยังเป็นตัวเลข ให้อ่านข้อมูลไปเพิ่่มใน DataTable                
                dr = dt.NewRow
                dr("รหัสสินค้า") = ws.Cells(row, 1).value
                dr("y") = ws.Cells(row, 2).value
                dr("z") = ws.Cells(row, 3).value
                dt.Rows.Add(dr)
            Else : HasData = False
            End If
            row += 1
            If row > 20 Then
                HasData = False
            End If
        Loop While (HasData)
        DataGridView1.DataSource = dt
        wb.Close()
    End Sub

    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Using b As SolidBrush = New SolidBrush(Me.DataGridView1.RowHeadersDefaultCellStyle.ForeColor)

            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), _
                                    Me.DataGridView1.DefaultCellStyle.Font, _
                                    b, e.RowBounds.Location.X + 5, _
                                    e.RowBounds.Location.Y + 5)
        End Using
        'เพิ่มตัวเลขที่ GridView
    End Sub

    Private Sub GetDataFromExcel()
        'Dim EXL As Excel.Application()
        'Dim CData As Excel.Range
        'Dim WSheet As Excel.Worksheet()


        'Me.TextBox1.Clear()
        'Dim WSheet As New Excel.Worksheet()
        ''Make excel file for test name is TEST.XLS (ensure that have Sheet1)
        'WSheet = EXL.Workbooks.Open("C:\TEST.XLS").Worksheets.Item("Sheet1")
        ''Define range of excel data ex.A1:Z1
        'EXL.Range("A2:E3").Select()
        'Dim CData As Excel.Range
        'CData = EXL.Selection
        'Dim iCol, iRow As Integer
        ''Begin loop for get data from excel to TextBox1.Text
        'For iCol = 1 To 5
        '    For iRow = 1 To 2
        '        TextBox1.Text = TextBox1.Text & _
        '        CData(iRow, iCol).value & vbTab
        '    Next
        '    TextBox1.Text = TextBox1.Text & vbCrLf
        'Next
        'EXL.Workbooks.Close()


    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class