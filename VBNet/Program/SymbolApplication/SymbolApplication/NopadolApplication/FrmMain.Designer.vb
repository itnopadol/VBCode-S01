<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FrmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private mainMenu1 As System.Windows.Forms.MainMenu

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TBBarCode = New System.Windows.Forms.TextBox
        Me.TBQty = New System.Windows.Forms.TextBox
        Me.BTNSave = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.ListViewBarCode = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.BTNAdd = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgName = New System.Windows.Forms.DataGrid
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(7, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 15)
        Me.Label2.Text = "BarCode :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(7, 61)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 15)
        Me.Label4.Text = "Qty :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBBarCode
        '
        Me.TBBarCode.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.TBBarCode.Location = New System.Drawing.Point(69, 34)
        Me.TBBarCode.Name = "TBBarCode"
        Me.TBBarCode.Size = New System.Drawing.Size(81, 18)
        Me.TBBarCode.TabIndex = 10
        '
        'TBQty
        '
        Me.TBQty.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.TBQty.Location = New System.Drawing.Point(69, 58)
        Me.TBQty.Name = "TBQty"
        Me.TBQty.Size = New System.Drawing.Size(81, 18)
        Me.TBQty.TabIndex = 13
        '
        'BTNSave
        '
        Me.BTNSave.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.BTNSave.Location = New System.Drawing.Point(69, 236)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(81, 20)
        Me.BTNSave.TabIndex = 14
        Me.BTNSave.Text = "Save"
        '
        'BTNExit
        '
        Me.BTNExit.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.BTNExit.Location = New System.Drawing.Point(156, 236)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(81, 20)
        Me.BTNExit.TabIndex = 15
        Me.BTNExit.Text = "Exit"
        '
        'ListViewBarCode
        '
        Me.ListViewBarCode.Columns.Add(Me.ColumnHeader1)
        Me.ListViewBarCode.Columns.Add(Me.ColumnHeader2)
        Me.ListViewBarCode.Columns.Add(Me.ColumnHeader3)
        Me.ListViewBarCode.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.ListViewBarCode.FullRowSelect = True
        Me.ListViewBarCode.Location = New System.Drawing.Point(69, 106)
        Me.ListViewBarCode.Name = "ListViewBarCode"
        Me.ListViewBarCode.Size = New System.Drawing.Size(236, 32)
        Me.ListViewBarCode.TabIndex = 18
        Me.ListViewBarCode.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ID"
        Me.ColumnHeader1.Width = 35
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "BarCode"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Qty"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader3.Width = 65
        '
        'BTNAdd
        '
        Me.BTNAdd.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.BTNAdd.Location = New System.Drawing.Point(156, 58)
        Me.BTNAdd.Name = "BTNAdd"
        Me.BTNAdd.Size = New System.Drawing.Size(21, 18)
        Me.BTNAdd.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Underline)
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(69, 94)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 15)
        Me.Label1.Text = "BarCode Select List"
        '
        'dgName
        '
        Me.dgName.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.dgName.Location = New System.Drawing.Point(69, 144)
        Me.dgName.Name = "dgName"
        Me.dgName.Size = New System.Drawing.Size(236, 61)
        Me.dgName.TabIndex = 23
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.CornflowerBlue
        Me.ClientSize = New System.Drawing.Size(323, 275)
        Me.ControlBox = False
        Me.Controls.Add(Me.dgName)
        Me.Controls.Add(Me.BTNAdd)
        Me.Controls.Add(Me.ListViewBarCode)
        Me.Controls.Add(Me.BTNExit)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.TBQty)
        Me.Controls.Add(Me.TBBarCode)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.mainMenu1
        Me.Name = "FrmMain"
        Me.Text = "Application"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TBBarCode As System.Windows.Forms.TextBox
    Friend WithEvents TBQty As System.Windows.Forms.TextBox
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents ListViewBarCode As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNAdd As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgName As System.Windows.Forms.DataGrid

End Class
