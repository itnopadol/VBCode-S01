Imports System.Management
Imports System.Net.IPHostEntry

Public Class FormPTT

    Private Sub FormPTT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim i As String

        Dim mc As System.Management.ManagementClass
        Dim mo As ManagementObject
        mc = New ManagementClass("Win32_Network­AdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        For Each mo In moc
            If mo.Item("IPEnabled") = True Then
                MsgBox("MAC address " & mo.Item("MacAddress").ToString())
            End If
        Next

        Me.TBLastCheckIn.Text = Now
        Me.TB1.Text = My.Computer.Name
        Me.TB2.Text = My.Computer.Info.OSFullName
        Me.TB3.Text = mo.Item("MacAddress").ToString()
    End Sub







End Class