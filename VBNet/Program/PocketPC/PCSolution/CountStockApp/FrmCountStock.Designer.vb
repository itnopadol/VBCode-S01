<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FrmCountStock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCountStock))
        Me.TMBackground = New System.Windows.Forms.Timer
        Me.PNShow = New System.Windows.Forms.Panel
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PNShow.SuspendLayout()
        Me.SuspendLayout()
        '
        'TMBackground
        '
        Me.TMBackground.Enabled = True
        Me.TMBackground.Interval = 1000
        '
        'PNShow
        '
        Me.PNShow.Controls.Add(Me.PictureBox1)
        Me.PNShow.Location = New System.Drawing.Point(158, 3)
        Me.PNShow.Name = "PNShow"
        Me.PNShow.Size = New System.Drawing.Size(79, 314)
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(1, -3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(240, 231)
        '
        'Panel1
        '
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(240, 320)
        '
        'FrmCountStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 320)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PNShow)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FrmCountStock"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PNShow.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TMBackground As System.Windows.Forms.Timer
    Friend WithEvents PNShow As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel

End Class
