<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrintDWpoint
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
        Me.CryRPTdw = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CryRPTdw
        '
        Me.CryRPTdw.ActiveViewIndex = -1
        Me.CryRPTdw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CryRPTdw.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CryRPTdw.Location = New System.Drawing.Point(0, 0)
        Me.CryRPTdw.Name = "CryRPTdw"
        Me.CryRPTdw.SelectionFormula = ""
        Me.CryRPTdw.Size = New System.Drawing.Size(763, 649)
        Me.CryRPTdw.TabIndex = 0
        Me.CryRPTdw.ViewTimeSelectionFormula = ""
        '
        'frmPrintDWpoint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(763, 649)
        Me.Controls.Add(Me.CryRPTdw)
        Me.Name = "frmPrintDWpoint"
        Me.Text = "frmPrintDWpoint"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CryRPTdw As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
