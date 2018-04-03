<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormReportPaidCommNet
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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Crystal101 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'Crystal101
        '
        Me.Crystal101.ActiveViewIndex = -1
        Me.Crystal101.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crystal101.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Crystal101.Location = New System.Drawing.Point(0, 0)
        Me.Crystal101.Name = "Crystal101"
        Me.Crystal101.SelectionFormula = ""
        Me.Crystal101.Size = New System.Drawing.Size(1016, 734)
        Me.Crystal101.TabIndex = 1
        Me.Crystal101.ViewTimeSelectionFormula = ""
        '
        'FormReportPaidCommNet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.Crystal101)
        Me.Name = "FormReportPaidCommNet"
        Me.Text = "FormReportPaidCommNet"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Crystal101 As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
