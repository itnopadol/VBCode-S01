Imports System.IO
Imports Symbol
Imports Symbol.Barcode
Imports Symbol.Barcode.Reader
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.ComponentModel
Imports System.Windows.Forms
Imports vb6 = Microsoft.VisualBasic
Public Class Form1
    Dim vQuery As String

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim vMemProfit As String
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String


        If vb6.InStr(Me.TextBox1.Text, "@") > 0 Then
            vMemProfit = "S01"
            vBarCode = vb6.Left(Me.TextBox1.Text, vb6.Len(Me.TextBox1.Text) - 1)

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails_Cat '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
            End If

            Me.TextBox2.Text = vItemCode
            Me.TextBox3.Text = vItemName

            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BTNClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        FormMainApplication.Show()
        Me.Hide()
    End Sub
End Class