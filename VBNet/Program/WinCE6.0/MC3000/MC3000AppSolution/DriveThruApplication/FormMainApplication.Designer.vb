<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FormMainApplication
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMainApplication))
        Me.BTNCheckOut = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TBUserName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BTNBack = New System.Windows.Forms.Button
        Me.BTNPickup = New System.Windows.Forms.Button
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'BTNCheckOut
        '
        Me.BTNCheckOut.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCheckOut.Location = New System.Drawing.Point(14, 146)
        Me.BTNCheckOut.Name = "BTNCheckOut"
        Me.BTNCheckOut.Size = New System.Drawing.Size(100, 33)
        Me.BTNCheckOut.TabIndex = 23
        Me.BTNCheckOut.Text = "2.CheckOut"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(161, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 13)
        Me.Label2.Text = "ชื่อผู้ใช้งาน :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBUserName
        '
        Me.TBUserName.Enabled = False
        Me.TBUserName.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBUserName.Location = New System.Drawing.Point(236, 14)
        Me.TBUserName.Name = "TBUserName"
        Me.TBUserName.Size = New System.Drawing.Size(75, 19)
        Me.TBUserName.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 10.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(14, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.Text = "โปรแกรมต่าง ๆ"
        '
        'BTNBack
        '
        Me.BTNBack.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.BTNBack.Location = New System.Drawing.Point(211, 254)
        Me.BTNBack.Name = "BTNBack"
        Me.BTNBack.Size = New System.Drawing.Size(100, 33)
        Me.BTNBack.TabIndex = 21
        Me.BTNBack.Text = "ESC-กลับ"
        '
        'BTNPickup
        '
        Me.BTNPickup.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.BTNPickup.Location = New System.Drawing.Point(14, 110)
        Me.BTNPickup.Name = "BTNPickup"
        Me.BTNPickup.Size = New System.Drawing.Size(100, 33)
        Me.BTNPickup.TabIndex = 20
        Me.BTNPickup.Text = "1.PickupApp"
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(70, 113)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(249, 187)
        '
        'FormMainApplication
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(325, 300)
        Me.ControlBox = False
        Me.Controls.Add(Me.BTNCheckOut)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TBUserName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BTNBack)
        Me.Controls.Add(Me.BTNPickup)
        Me.Controls.Add(Me.PictureBox5)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormMainApplication"
        Me.Text = "FormMainApplication"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BTNCheckOut As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TBUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BTNBack As System.Windows.Forms.Button
    Friend WithEvents BTNPickup As System.Windows.Forms.Button
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
End Class
