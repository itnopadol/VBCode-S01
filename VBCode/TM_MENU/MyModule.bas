Attribute VB_Name = "MyModule"
Option Explicit

Global gConnection As New ADODB.Connection
Global gSoftwareLink() As String
Global gNumOfSoftware As Integer

Public Function FormatDataType(pValue, pFormat As String) As String
    Dim vReturnValue As String
    Select Case UCase(pFormat)
        Case "DATE"
            If IsDate(pValue) Then
                vReturnValue = Format(pValue, "Short Date")
            Else
                vReturnValue = ""
            End If
        Case "CURRENCY"
            If IsNumeric(pValue) Then
                vReturnValue = Format(SetNumber(pValue), "#,##0")
            Else
                vReturnValue = "0"
            End If
        Case "CURRENCY2"
            If IsNumeric(pValue) Then
                vReturnValue = Format(SetNumber(pValue), "#,##0.00")
            Else
                vReturnValue = "0.00"
            End If
        Case "CURRENCY3"
            If IsNumeric(pValue) Then
                vReturnValue = Format(SetNumber(pValue), "#,##0.000")
            Else
                vReturnValue = "0"
            End If
        Case "CURRENCY4"
            If IsNumeric(pValue) Then
                vReturnValue = Format(SetNumber(pValue), "#,##0.0000")
            Else
                vReturnValue = "0.00"
            End If
        Case "INTEGER"
            If IsNumeric(pValue) Then
                vReturnValue = Trim(Str(Format(SetNumber(pValue), "#0")))
            Else
                vReturnValue = "0"
            End If
        Case "NUMERIC"
            If IsNumeric(pValue) Then
                vReturnValue = Trim(Str(SetNumber(pValue)))
            Else
                vReturnValue = "0"
            End If
        Case "TRIM"
            vReturnValue = Trim(pValue)
        Case "TRIMCAPS"
            vReturnValue = Trim(UCase(pValue))
        Case Else
            vReturnValue = pValue
    End Select
    FormatDataType = vReturnValue
End Function


Public Function GetField(pRecordset As ADODB.Recordset, pFieldName)
    Dim vReturnValue
    
    Select Case pRecordset.Fields(pFieldName).Type
        Case 2, 3, 4, 5, 6, 17, 20, 72, 131 'Numeric
            If IsNull(pRecordset.Fields(pFieldName).Value) Then
                vReturnValue = 0
            Else
                vReturnValue = IIf(IsNull(pRecordset.Fields(pFieldName).Value), 0, pRecordset.Fields(pFieldName).Value)
            End If
        Case 129, 130, 200, 201, 202, 203 'Text
            If IsNull(pRecordset.Fields(pFieldName).Value) Then
                vReturnValue = ""
            Else
                vReturnValue = IIf(IsNull(pRecordset.Fields(pFieldName).Value), "", Trim(pRecordset.Fields(pFieldName).Value))
            End If
        Case 135 'Date/Time
            If IsNull(pRecordset.Fields(pFieldName).Value) Then
                vReturnValue = ""
            Else
                vReturnValue = IIf(IsNull(pRecordset.Fields(pFieldName).Value), "", Format(pRecordset.Fields(pFieldName).Value, "Short Date"))
            End If
    End Select
    GetField = vReturnValue
End Function

Public Function OpenTable(pConnection As ADODB.Connection, pRecordset As ADODB.Recordset, _
    pQuery As String) As Long
    
    pRecordset.CursorLocation = adUseClient
    pRecordset.Open pQuery, pConnection, adOpenDynamic, adLockOptimistic
    OpenTable = pRecordset.RecordCount
End Function

Public Sub PutField(pRecordset As ADODB.Recordset, pFieldName, pValue)
    Dim vReturnValue
    Dim vDate As Date
    
    Select Case pRecordset.Fields(pFieldName).Type
        Case 2, 3, 4, 5, 6, 17, 20, 72, 131 'Numeric
            If IsNumeric(pValue) Then
                vReturnValue = SetNumber(pValue)
            Else
                vReturnValue = 0
            End If
        Case 129, 130, 200, 201, 202, 203 'Text
            vReturnValue = Trim(pValue)
        Case 135 'Date/Time
            vDate = Format(pValue, "Short Date")
            vReturnValue = vDate
    End Select
    pRecordset.Fields(pFieldName).Value = vReturnValue
End Sub

Public Function SetNumber(pNumber)
    Dim vStop As Integer
    Dim vValue As String
    Dim vReturnValue As String
    Dim i As Integer
    Dim vChar As String
    
    vReturnValue = ""
    vValue = pNumber
    vValue = Trim(vValue)
    vStop = Len(vValue)
    For i = 1 To vStop
        vChar = Mid(vValue, i, 1)
        If IsNumeric(vChar) Or vChar = "." Or vChar = "-" Then
            vReturnValue = vReturnValue & vChar
        End If
    Next
    SetNumber = Val(vReturnValue)
End Function
