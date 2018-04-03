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

        Me.TBBarCode.Focus()

    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        Me.TBGetBarCode.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBITemName.Text = ""
        Me.TBItemStatus.Text = ""
        Me.TBItemType.Text = ""
        Me.TBPriceType.Text = ""
        Me.TBRVDate.Text = ""
        Me.TBRVNo.Text = ""
        Me.TBRVUserID.Text = ""
        Me.TBSalePrice1.Text = ""
        Me.TBSalePrice2.Text = ""
        Me.TBShelfID.Text = ""
        Me.TBStkUnit.Text = ""
        Me.TBUnit.Text = ""
        Me.TBBarCode.Focus()
    End Sub

    Public Sub ClearData()
        On Error Resume Next

        Me.TBGetBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBITemName.Text = ""
        Me.TBItemStatus.Text = ""
        Me.TBItemType.Text = ""
        Me.TBPriceType.Text = ""
        Me.TBRVDate.Text = ""
        Me.TBRVNo.Text = ""
        Me.TBRVUserID.Text = ""
        Me.TBSalePrice1.Text = ""
        Me.TBSalePrice2.Text = ""
        Me.TBShelfID.Text = ""
        Me.TBStkUnit.Text = ""
        Me.TBUnit.Text = ""
    End Sub

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vBarCode As String
        Dim vSalePrice1 As Double
        Dim vSalePrice2 As Double

        On Error Resume Next

        If e.KeyCode = Keys.Enter And Me.TBBarCode.Text <> "" Then

            vBarCode = Me.TBBarCode.Text

            vQuery = "exec dbo.USP_NP_SearchBarCodeCheckScan '" & vBarCode & "'"
            Call vGetData1(vMemProfit, vQuery)
            If pds1.Tables(0).Rows.Count > 0 Then

                vSalePrice1 = pds1.Tables(0).Rows(0)("saleprice1").ToString
                vSalePrice2 = pds1.Tables(0).Rows(0)("saleprice2").ToString

                Me.TBGetBarCode.Text = pds1.Tables(0).Rows(0)("barcode").ToString
                Me.TBItemCode.Text = pds1.Tables(0).Rows(0)("code").ToString
                Me.TBITemName.Text = pds1.Tables(0).Rows(0)("name1").ToString
                Me.TBItemStatus.Text = pds1.Tables(0).Rows(0)("itemstatus").ToString
                Me.TBItemType.Text = pds1.Tables(0).Rows(0)("itemtype").ToString
                Me.TBPriceType.Text = pds1.Tables(0).Rows(0)("remark").ToString
                Me.TBRVDate.Text = "" 'pds1.Tables(0).Rows(0)("shelf").ToString
                Me.TBRVNo.Text = pds1.Tables(0).Rows(0)("rvno").ToString
                Me.TBRVUserID.Text = pds1.Tables(0).Rows(0)("rvuserid").ToString
                Me.TBSalePrice1.Text = Format(vSalePrice1, "##,##0.00")
                Me.TBSalePrice2.Text = Format(vSalePrice2, "##,##0.00")
                Me.TBShelfID.Text = pds1.Tables(0).Rows(0)("shelfid1").ToString
                Me.TBStkUnit.Text = pds1.Tables(0).Rows(0)("defstkunitcode").ToString
                Me.TBUnit.Text = pds1.Tables(0).Rows(0)("unitcode").ToString
            Else
                Me.TBGetBarCode.Text = ""
                Me.TBItemCode.Text = ""
                Me.TBITemName.Text = ""
                Me.TBItemStatus.Text = ""
                Me.TBItemType.Text = ""
                Me.TBPriceType.Text = ""
                Me.TBRVDate.Text = ""
                Me.TBRVNo.Text = ""
                Me.TBRVUserID.Text = ""
                Me.TBSalePrice1.Text = ""
                Me.TBSalePrice2.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBStkUnit.Text = ""
                Me.TBUnit.Text = ""
                Call ClearData()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        Dim vBarCode As String
        Dim vSalePrice1 As Double
        Dim vSalePrice2 As Double

        On Error Resume Next

        If vb6.InStr(Me.TBBarCode.Text, "@") > 0 Then

            vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)
            Me.TBBarCode.Text = vBarCode
            vBarCode = Me.TBBarCode.Text

            vQuery = "exec dbo.USP_NP_SearchBarCodeCheckScan '" & vBarCode & "'"
            Call vGetData1(vMemProfit, vQuery)
            If pds1.Tables(0).Rows.Count > 0 Then

                vSalePrice1 = pds1.Tables(0).Rows(0)("saleprice1").ToString
                vSalePrice2 = pds1.Tables(0).Rows(0)("saleprice2").ToString

                Me.TBGetBarCode.Text = pds1.Tables(0).Rows(0)("barcode").ToString
                Me.TBItemCode.Text = pds1.Tables(0).Rows(0)("code").ToString
                Me.TBITemName.Text = pds1.Tables(0).Rows(0)("name1").ToString
                Me.TBItemStatus.Text = pds1.Tables(0).Rows(0)("itemstatus").ToString
                Me.TBItemType.Text = pds1.Tables(0).Rows(0)("itemtype").ToString
                Me.TBPriceType.Text = pds1.Tables(0).Rows(0)("remark").ToString
                Me.TBRVDate.Text = "" 'pds1.Tables(0).Rows(0)("shelf").ToString
                Me.TBRVNo.Text = pds1.Tables(0).Rows(0)("rvno").ToString
                Me.TBRVUserID.Text = pds1.Tables(0).Rows(0)("rvuserid").ToString
                Me.TBSalePrice1.Text = Format(vSalePrice1, "##,##0.00")
                Me.TBSalePrice2.Text = Format(vSalePrice2, "##,##0.00")
                Me.TBShelfID.Text = pds1.Tables(0).Rows(0)("shelfid1").ToString
                Me.TBStkUnit.Text = pds1.Tables(0).Rows(0)("defstkunitcode").ToString
                Me.TBUnit.Text = pds1.Tables(0).Rows(0)("unitcode").ToString
            Else
                Me.TBGetBarCode.Text = ""
                Me.TBItemCode.Text = ""
                Me.TBITemName.Text = ""
                Me.TBItemStatus.Text = ""
                Me.TBItemType.Text = ""
                Me.TBPriceType.Text = ""
                Me.TBRVDate.Text = ""
                Me.TBRVNo.Text = ""
                Me.TBRVUserID.Text = ""
                Me.TBSalePrice1.Text = ""
                Me.TBSalePrice2.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBStkUnit.Text = ""
                Me.TBUnit.Text = ""
                Call ClearScreen()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If
        End If

        If Me.TBBarCode.Text = "" Then
            Call ClearData()
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub
End Class