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

Public Class FormTestScan
    Dim vQuery As String
    Dim vMemDocDate As String

    Private Sub FormCheckShelf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim n As Integer
        Dim vShelf As String

        Me.TBBarCode.Focus()

        'vQuery = "exec dbo.USP_MB_ShelfPlanScanBarcode 'S01','01/01/2011'"
        'Call vGetData1(vMemProfit, vQuery)
        'If pds1.Tables(0).Rows.Count > 0 Then

        '    For n = 0 To pds1.Tables(0).Rows.Count - 1
        '        vShelf = pds1.Tables(0).Rows(n)("shelf").ToString

        '        Dim listItem As New ListViewItem(vShelf)
        '        Me.ListViewShelf.Items.Add(listItem)

        '        Me.ListViewShelf.SmallImageList = Me.ImageList1
        '        Me.ListViewShelf.Items(n).ImageIndex = 1
        '    Next
        'End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        MsgBox(e.KeyCode)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged

    End Sub
End Class