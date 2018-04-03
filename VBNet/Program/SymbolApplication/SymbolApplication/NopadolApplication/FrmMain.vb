Imports System.Data
Imports System.Data.SqlServerCe
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.ComponentModel
Imports System.Windows.Forms

Public Class FrmMain

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call BindDataGrid()
        Me.TBBarCode.Focus()
    End Sub

    Private Sub BTNAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAdd.Click
        Dim i As Integer

        i = 1
        Dim listItem As New ListViewItem(i)
        listItem.SubItems.Add(Me.TBBarCode.Text)
        listItem.SubItems.Add(Me.TBQty.Text)
        Me.ListViewBarCode.Items.Add(listItem)

    End Sub

    Private Sub BindDataGrid()

        Dim myConnection As SqlCeConnection
        Dim dt As New DataTable
        Dim Adapter As SqlCeDataAdapter

        myConnection = New SqlCeConnection("Data Source ="(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase) + "\NPDB.sdf;"))
        myConnection.Open()

        Dim myCommand As SqlCeCommand = myConnection.CreateCommand()
        myCommand.CommandText = "SELECT [id], [name], [email] FROM [mytable]"
        myCommand.CommandType = CommandType.Text

        Adapter = New SqlCeDataAdapter(myCommand)
        Adapter.Fill(dt)

        myConnection.Close()

        Dim tableStyle As New DataGridTableStyle()

        tableStyle.MappingName = dt.TableName

        Dim column As New DataGridTextBoxColumn()
        column.MappingName = "id"
        column.HeaderText = "ID"
        column.Width = 30

        tableStyle.GridColumnStyles.Add(column)

        column = New DataGridTextBoxColumn()
        column.MappingName = "name"
        column.HeaderText = "Name"
        column.Width = 40
        tableStyle.GridColumnStyles.Add(column)

        column = New DataGridTextBoxColumn()
        column.Width = 70
        column.MappingName = "email"
        column.HeaderText = "Email"
        tableStyle.GridColumnStyles.Add(column)

        Me.dgName.DataSource = dt
        Me.dgName.TableStyles.Clear()
        Me.dgName.TableStyles.Add(tableStyle)

        dt = Nothing

    End Sub
End Class
