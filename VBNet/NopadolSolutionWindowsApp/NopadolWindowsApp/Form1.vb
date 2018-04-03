Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    Private dataGridView As New DataGridView()
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCIndex As Integer
    Dim vRIndex As Integer

    <STAThreadAttribute()> _
    Public Shared Sub Main()
        Application.Run(New Form1())
    End Sub

    Public Sub New1()
        Me.DataGridView1.Dock = DockStyle.Fill
        Me.Controls.Add(Me.DataGridView1)
        Me.Text = "DataGridView calendar column demo"
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) _
        Handles Me.Load

        Call InitializeDataBase1()
        'Dim col As New CalendarColumn
        'Me.DataGridView1.Columns.Add(col)
        'Me.DataGridView1.RowCount = 1

        ''For Each row In Me.dataGridView1.Rows
        ''    row.Cells(0).Value = DateTime.Now
        ''Next row



        ''Dim colgcus_ch As New DataGridViewCheckBoxColumn()
        ''Dim colgcus_cmb As New DataGridViewComboBoxColumn

        ''Dim colgcus_txt As New DataGridViewTextBoxColumn
        ''Dim colgcus_txt1 As New DataGridViewTextBoxColumn
        ''Dim colgcus_txt2 As New DataGridViewTextBoxColumn
        ''Dim colgcus_txt3 As New DataGridViewTextBoxColumn


        ''DataGridView1.Columns.Add(colgcus_txt)
        ''DataGridView1.Columns.Add(colgcus_txt1)
        ''DataGridView1.Columns.Add(colgcus_txt2)
        ''DataGridView1.Columns.Add(colgcus_txt3)

        'Dim col1 As New CalendarColumn

        'Me.DataGridView1.Columns.Add(col1)
        'Me.DataGridView1.RowCount = 1



        'vQuery = "exec dbo.usp_ps_brandlist"
        'da = New SqlDataAdapter(vQuery, vConnection)
        'ds = New DataSet
        'da.Fill(ds, "Brand")
        'dt = ds.Tables("Brand")

        ''Me.Column1.DataSource = dt
        ''Me.Column1.DisplayMember = "Brand"
        ''Me.Column1.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        ''Me.Column1.DisplayStyleForCurrentCellOnly = True
        ''Me.Column1.HeaderText = "Column1"
        ''Me.Column1.Name = "Column1"
        ''Me.Column1.ValueMember = "brand"


        ' ''DataGridView1.Columns.Add(colgcus_cmb)


        ''Dim row1 As DataGridViewRow
        ''For Each row In Me.DataGridView1.Rows
        ''    row1.Cells(0).Value = DateTime.Now
        ''Next row

        ''datagridview1.Columns("Cost").DefaultCellStyle.Format = "#,##0.00" จัด format ที่มีทศนิยม

        ''DataGridView1.Columns(0).DataGridView.DataSource = dt
        ''DataGridView1.Columns(0).Displayed = True
        ''DataGridView1.Columns(0).
        ''Me.Column1.DataSource = Me.CategoryBindingSource
        ''Me.Column1.DisplayMember = "CategoryName"
        ''Me.Column1.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        ''Me.Column1.DisplayStyleForCurrentCellOnly = True
        ''Me.Column1.HeaderText = "Column1"
        ''Me.Column1.Name = "Column1"
        ''Me.Column1.ValueMember = "CategoryID"
        ''Call AddComboBoxColumns()



    End Sub
    Enum ColumnName
        EmployeeId
        LastName
        FirstName
        Title
        TitleOfCourtesy
        BirthDate
        HireDate
        Address
        City
        Region
        PostalCode
        Country
        HomePhone
        Extension
        Photo
        Notes
        ReportsTo
        PhotoPath
        OutOfOffice
    End Enum

    Private Sub AddComboBoxColumns()
        Dim comboboxColumn As New DataGridViewComboBoxColumn()
        comboboxColumn = CreateComboBoxColumn()
        SetAlternateChoicesUsingDataSource(comboboxColumn)
        comboboxColumn.HeaderText = _
            "TitleOfCourtesy"
        DataGridView1.Columns.Insert(0, comboboxColumn)

        comboboxColumn = CreateComboBoxColumn()
        SetAlternateChoicesUsingItems(comboboxColumn)
        comboboxColumn.HeaderText = _
            "TitleOfCourtesy"
        ' Tack this example column onto the end.
        DataGridView1.Columns.Add(comboboxColumn)
    End Sub
    Private Function CreateComboBoxColumn() _
    As DataGridViewComboBoxColumn
        Dim column As New DataGridViewComboBoxColumn()

        With column
            .DataPropertyName = ColumnName.TitleOfCourtesy.ToString()
            .HeaderText = ColumnName.TitleOfCourtesy.ToString()
            .DropDownWidth = 160
            .Width = 90
            .MaxDropDownItems = 3
            .FlatStyle = FlatStyle.Flat
        End With
        Return column
    End Function

    Private Sub SetAlternateChoicesUsingDataSource( _
        ByRef comboboxColumn As DataGridViewComboBoxColumn)
        With comboboxColumn
            .DataSource = RetrieveAlternativeTitles()
            .ValueMember = ColumnName.TitleOfCourtesy.ToString()
            .DisplayMember = .ValueMember
        End With
    End Sub
    Private Shared Sub SetAlternateChoicesUsingItems( _
        ByRef comboboxColumn As DataGridViewComboBoxColumn)

        With comboboxColumn
            .Items.AddRange(New String() _
                    {"Mr.", "Ms.", "Mrs.", "Dr."})
        End With
    End Sub
    Private Function RetrieveAlternativeTitles() As DataTable
        Return Populate("SELECT distinct TitleOfCourtesy FROM Employees")
    End Function


    Private Function Populate(ByVal sqlCommand As String) As DataTable
        Dim northwindConnection As New SqlConnection(connectionString)
        northwindConnection.Open()

        Dim command As New SqlCommand(sqlCommand, _
            northwindConnection)
        Dim adapter As New SqlDataAdapter()
        adapter.SelectCommand = command
        Dim table As New DataTable()
        table.Locale = System.Globalization.CultureInfo.InvariantCulture
        adapter.Fill(table)

        Return table
    End Function
    Private connectionString As String = _
            "Integrated Security=SSPI;Persist Security Info=False;" _
            & "Initial Catalog=Northwind;Data Source=nebula"

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

        'MsgBox(DataGridView1.Columns(e.ColumnIndex).Name)
        vCIndex = e.ColumnIndex
        vRIndex = e.RowIndex

        If DataGridView1.Columns(e.ColumnIndex).Name = "Column2" Then

            'Me.GBSearchCategoryCode.Visible = True
            DataGridView1.Rows(vRIndex).DataGridView.Rows.Clear()

            Dim drv As DataRowView = DataGridView1.CurrentRow.DataBoundItem

            If (drv IsNot Nothing) Then
                drv.Row.Delete()
            End If


        End If

    End Sub


    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Using b As SolidBrush = New SolidBrush(Me.dataGridView.RowHeadersDefaultCellStyle.ForeColor)

            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), _
                                    Me.dataGridView.DefaultCellStyle.Font, _
                                    b, e.RowBounds.Location.X + 5, _
                                    e.RowBounds.Location.Y + 5)
        End Using
        'เพิ่มตัวเลขที่ GridView
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Column1.DataGridView(vCIndex - 1, vRIndex).Value = "1000"
        Me.Column3.DataGridView(vCIndex + 1, vRIndex).Selected = True
        GBSearchCategoryCode.Visible = False
    End Sub


    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

        Dim drv As DataRowView = DataGridView1.CurrentRow.DataBoundItem
        If (drv IsNot Nothing) Then
            drv.Row.Delete()
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub
End Class



'Public Class Form1
'    Dim ds As DataSet
'    Dim da As SqlDataAdapter
'    Dim dt As DataTable
'    Dim vQuery As String
'    Dim vCMD As SqlCommand
'    Dim vDocno As String

'    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        Call InitializeDataBase()
'        GetBrand()
'    End Sub
'    Public Sub GetBrand()
'        Dim colgcus_ch As New DataGridViewCheckBoxColumn()
'        Dim colgcus_cmb As New DataGridViewComboBoxColumn
'        Dim colgcus_txt As New DataGridViewTextBoxColumn

'        colgcus_ch.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

'        colgcus_ch.DataPropertyName = "AAAAA"

'        colgcus_ch.HeaderText = "AAAAA"

'        colgcus_cmb.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

'        colgcus_cmb.HeaderText = "BBBB"

'        colgcus_txt.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

'        colgcus_txt.HeaderText = "CCCCC"

'        DataGridView1.Columns.Add(colgcus_ch)
'        DataGridView1.RowCount = 5

'        DataGridView1.Columns.Add(colgcus_cmb)
'        DataGridView1.RowCount = 5

'        DataGridView1.Columns.Add(colgcus_txt)
'        DataGridView1.RowCount = 5
'    End Sub

'    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

'    End Sub

'    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

'    End Sub
'End Class