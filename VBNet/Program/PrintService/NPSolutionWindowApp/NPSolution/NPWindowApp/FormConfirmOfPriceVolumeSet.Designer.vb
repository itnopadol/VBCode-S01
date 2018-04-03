<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormConfirmOfPriceVolumeSet
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
        Me.pnSetConfig = New System.Windows.Forms.Panel
        Me.btnSelectProduct = New System.Windows.Forms.Button
        Me.pnSetConfig.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnSetConfig
        '
        Me.pnSetConfig.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnSetConfig.Controls.Add(Me.btnSelectProduct)
        Me.pnSetConfig.Location = New System.Drawing.Point(12, 76)
        Me.pnSetConfig.Name = "pnSetConfig"
        Me.pnSetConfig.Size = New System.Drawing.Size(1025, 700)
        Me.pnSetConfig.TabIndex = 3
        '
        'btnSelectProduct
        '
        Me.btnSelectProduct.Image = Global.NPWindowApp.My.Resources.Resources.icon_16_checkin
        Me.btnSelectProduct.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSelectProduct.Location = New System.Drawing.Point(761, 660)
        Me.btnSelectProduct.Name = "btnSelectProduct"
        Me.btnSelectProduct.Size = New System.Drawing.Size(104, 37)
        Me.btnSelectProduct.TabIndex = 7
        Me.btnSelectProduct.Text = "เลือกสินค้า"
        Me.btnSelectProduct.UseVisualStyleBackColor = True
        '
        'FormConfirmOfPriceVolumeSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1007, 680)
        Me.Controls.Add(Me.pnSetConfig)
        Me.Name = "FormConfirmOfPriceVolumeSet"
        Me.Text = "FormConfirmOfPriceVolumeSet"
        Me.pnSetConfig.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnSetConfig As System.Windows.Forms.Panel
    Friend WithEvents btnSelectProduct As System.Windows.Forms.Button
End Class
