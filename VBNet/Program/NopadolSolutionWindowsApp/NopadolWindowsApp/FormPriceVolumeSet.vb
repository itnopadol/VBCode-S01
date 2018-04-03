Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Public Class FormPriceVolumeSet
    Dim Qrystr As String

    Private Sub FormPriceVolumeSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized
    End Sub

    


    Private Sub btnProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class