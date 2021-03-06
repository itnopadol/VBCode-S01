Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Module Module1
    Public vConnectionString As String
    Public vConnection As SqlConnection
    Public vCon As SqlConnection
    Public vConnectionStringBPlus As String
    Public vConnectionBPlus As SqlConnection

    Public da As SqlDataAdapter
    Public ds As DataSet
    Public vdt As DataTable
    Public dt As New DataTable
    Public vUserID As String
    Public vPassword As String
    Public vDepartment As String
    Public vLevelID As Integer

    '---------------------------------
    'Valiable for Pric Volume Set
    Public publicFdocno As String
    Public PublicDocStatus As String
    Public sv As String
    Public frmPriceVolumeSet As frmPriceVolumeSet
    Public FormApproveVolumeSet As FormApproveVolumeSet
    Public pvDocno As String
    Public pvLV2 As String
    Public pvLV3 As String
    Public pvVM2 As String
    Public pvVM3 As String
    Public pvDC2 As String
    Public pvDC3 As String
    Public pvSMP2 As String
    Public pvSMP3 As String
    Public lpStatus As Integer

    '---------------------------------


    Public vIncentiveDocNo As String '���Ţ������ Incentive
    Public vPriceStructureDocNo As String '���Ţ������ Incentive



    Public Sub InitializeDataBase()
        vConnectionString = "Persist Security Info = False;User ID='" & vUserID & "';Password='" & vPassword & "';Data Source = Nebula;Initial Catalog = BCNP"
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()
    End Sub

    Public Sub InitializeDataBase1()
        vConnectionString = "Persist Security Info = False;User ID='panuvich';Password='thaikom$';Data Source = Nebula;Initial Catalog = BCNP"
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()
    End Sub

    Public Sub InitializeDataBaseBPlus()
        vConnectionStringBPlus = "Persist Security Info = False;User ID='vbuser';Password='132';Data Source = Nebula;Initial Catalog = BCNP"
        vConnectionBPlus = New SqlConnection(vConnectionStringBPlus)
        vConnectionBPlus.Open()
    End Sub

    Public Sub ChekAuthorityAccess()
        Dim vQuery As String

        Call InitializeDataBaseBPlus()
        vQuery = "select department ,levelid from bcnp.dbo.vw_np_UserAutorityProgram where code = '" & vUserID & "' "
        da = New SqlDataAdapter(vQuery, vConnectionBPlus)
        ds = New DataSet
        da.Fill(ds, "Autority")
        dt = ds.Tables("Autority")

        If dt.Rows.Count > 0 Then
            vDepartment = Trim(dt.Rows(0).Item("department"))
            vLevelID = Trim(dt.Rows(0).Item("levelid"))
        Else
            MsgBox("����� User " & vUserID & " �����к� ��سҵ�Ǩ�ͺ ", MsgBoxStyle.Critical, "Send Error")
        End If
    End Sub
    '--------------------------------------
    Public Class CalendarColumn
        Inherits DataGridViewColumn

        Public Sub New()
            MyBase.New(New CalendarCell())
        End Sub

        Public Overrides Property CellTemplate() As DataGridViewCell
            Get
                Return MyBase.CellTemplate
            End Get
            Set(ByVal value As DataGridViewCell)

                ' Ensure that the cell used for the template is a CalendarCell.
                If (value IsNot Nothing) AndAlso _
                    Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) _
                    Then
                    Throw New InvalidCastException("Must be a CalendarCell")
                End If
                MyBase.CellTemplate = value

            End Set
        End Property

    End Class

    Public Class CalendarCell
        Inherits DataGridViewTextBoxCell

        Public Sub New()
            ' Use the short date format.
            Me.Style.Format = "d"
        End Sub

        Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
            ByVal initialFormattedValue As Object, _
            ByVal dataGridViewCellStyle As DataGridViewCellStyle)

            ' Set the value of the editing control to the current cell value.
            MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, _
                dataGridViewCellStyle)

            Dim ctl As CalendarEditingControl = _
                CType(DataGridView.EditingControl, CalendarEditingControl)
            ctl.Value = CType(Me.Value, DateTime)

        End Sub

        Public Overrides ReadOnly Property EditType() As Type
            Get
                ' Return the type of the editing contol that CalendarCell uses.
                Return GetType(CalendarEditingControl)
            End Get
        End Property

        Public Overrides ReadOnly Property ValueType() As Type
            Get
                ' Return the type of the value that CalendarCell contains.
                Return GetType(DateTime)
            End Get
        End Property

        Public Overrides ReadOnly Property DefaultNewRowValue() As Object
            Get
                ' Use the current date and time as the default value.
                Return DateTime.Now
            End Get
        End Property

    End Class

    Class CalendarEditingControl
        Inherits DateTimePicker
        Implements IDataGridViewEditingControl

        Private dataGridViewControl As DataGridView
        Private valueIsChanged As Boolean = False
        Private rowIndexNum As Integer

        Public Sub New()
            Me.Format = DateTimePickerFormat.Short
        End Sub

        Public Property EditingControlFormattedValue() As Object _
            Implements IDataGridViewEditingControl.EditingControlFormattedValue

            Get
                Return Me.Value.ToShortDateString()
            End Get

            Set(ByVal value As Object)
                If TypeOf value Is String Then
                    Me.Value = DateTime.Parse(CStr(value))
                End If
            End Set

        End Property

        Public Function GetEditingControlFormattedValue(ByVal context _
            As DataGridViewDataErrorContexts) As Object _
            Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

            Return Me.Value.ToShortDateString()

        End Function

        Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As _
            DataGridViewCellStyle) _
            Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

            Me.Font = dataGridViewCellStyle.Font
            Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
            Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

        End Sub

        Public Property EditingControlRowIndex() As Integer _
            Implements IDataGridViewEditingControl.EditingControlRowIndex

            Get
                Return rowIndexNum
            End Get
            Set(ByVal value As Integer)
                rowIndexNum = value
            End Set

        End Property

        Public Function EditingControlWantsInputKey(ByVal key As Keys, _
            ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
            Implements IDataGridViewEditingControl.EditingControlWantsInputKey

            ' Let the DateTimePicker handle the keys listed.
            Select Case key And Keys.KeyCode
                Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
                    Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

                    Return True

                Case Else
                    Return False
            End Select

        End Function

        Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
            Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

            ' No preparation needs to be done.

        End Sub

        Public ReadOnly Property RepositionEditingControlOnValueChange() _
            As Boolean Implements _
            IDataGridViewEditingControl.RepositionEditingControlOnValueChange

            Get
                Return False
            End Get

        End Property

        Public Property EditingControlDataGridView() As DataGridView _
            Implements IDataGridViewEditingControl.EditingControlDataGridView

            Get
                Return dataGridViewControl
            End Get
            Set(ByVal value As DataGridView)
                dataGridViewControl = value
            End Set

        End Property

        Public Property EditingControlValueChanged() As Boolean _
            Implements IDataGridViewEditingControl.EditingControlValueChanged

            Get
                Return valueIsChanged
            End Get
            Set(ByVal value As Boolean)
                valueIsChanged = value
            End Set

        End Property

        Public ReadOnly Property EditingControlCursor() As Cursor _
            Implements IDataGridViewEditingControl.EditingPanelCursor

            Get
                Return MyBase.Cursor
            End Get

        End Property

        Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)

            ' Notify the DataGridView that the contents of the cell have changed.
            valueIsChanged = True
            Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
            MyBase.OnValueChanged(eventargs)

        End Sub

    End Class


    Public Class Mydatagridview

        Inherits DataGridView

        Protected Overrides Function ProcessDialogKey(ByVal keydata As Keys) As Boolean

            Dim key As Keys = keydata And Keys.KeyCode

            If key = Keys.Enter Then

                Return Me.ProcessTabKey(keydata)

            End If

            Return MyBase.ProcessDialogKey(keydata)

        End Function


        Protected Overrides Function ProcessdatagridviewKey(ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean

            If e.KeyCode = Keys.Enter Then

                Return Me.ProcessTabKey(e.KeyData)

            End If

            Return MyBase.ProcessDataGridViewKey(e)

        End Function

    End Class
End Module
