Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.VisualBasic


Public Class Form2
    Dim btnSelector As Button = New Button
    Dim pCase As Integer

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CreateSeletor()
        'btnSelector.FlatStyle = FlatStyle.Flat
        'btnSelector.FlatAppearance.BorderSize = 0
        'btnSelector.Size = New Size(19, 19)
        'btnSelector.ImageAlign = ContentAlignment.MiddleCenter
        'btnSelector.FlatAppearance.MouseDownBackColor = Color.Transparent
        'btnSelector.FlatAppearance.MouseOverBackColor = Color.Transparent
        'btnSelector.BackColor = Color.Transparent
        'btnSelector.Image = DataGridviewbutton.Properties.Resources.search
        'dataGridView1.Controls.Add(btnSelector)
        'btnSelector.Hide()
        'btnSelector.Click += this.SelectorClick

    End Sub


    Private Sub SelectorClick(ByVal sender As System.Object, ByVal e As System.EventArgs)



        'pCase = dataGridView1.CurrentCell.ColumnIndex

        'Int(pRowFind = dataGridView1.CurrentRow.Index)

        'Switch(pCase)


    End Sub


    Private Sub dataGridView1_CellEnter(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'If e.ColumnIndex = 1 Then
        '    var(Loc() = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False))
        '    var(Wid = dataGridView1.CurrentCell.Size.Width)
        '    btnSelector.Location = New Point(Loc.X - 20 + Wid, Loc.Y)
        '    btnSelector.Show()
        'End If

    End Sub


    Private Sub dataGridView1_CellLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If btnSelector.Focused <> True Then

            btnSelector.Hide()
        End If
    End Sub


End Class