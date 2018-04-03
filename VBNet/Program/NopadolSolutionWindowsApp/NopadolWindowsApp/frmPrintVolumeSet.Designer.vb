<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrintVolumeSet
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
        Me.crtVW01 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'crtVW01
        '
        Me.crtVW01.ActiveViewIndex = -1
        Me.crtVW01.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVW01.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVW01.Location = New System.Drawing.Point(0, 0)
        Me.crtVW01.Name = "crtVW01"
        Me.crtVW01.SelectionFormula = ""
        Me.crtVW01.Size = New System.Drawing.Size(901, 644)
        Me.crtVW01.TabIndex = 0
        Me.crtVW01.ViewTimeSelectionFormula = ""
        '
        'frmPrintVolumeSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(901, 644)
        Me.Controls.Add(Me.crtVW01)
        Me.Name = "frmPrintVolumeSet"
        Me.Text = "พิมพ์เอกสารกำหนดราคาตาม  Volume"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crtVW01 As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
