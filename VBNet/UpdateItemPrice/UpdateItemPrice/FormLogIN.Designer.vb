<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormLogIN
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextUser = New System.Windows.Forms.TextBox
        Me.TextPassword = New System.Windows.Forms.TextBox
        Me.BTNOK = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(116, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ชื่อเข้าโปรแกรม"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(116, 90)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "รหัสผ่าน"
        '
        'TextUser
        '
        Me.TextUser.Location = New System.Drawing.Point(208, 51)
        Me.TextUser.Name = "TextUser"
        Me.TextUser.Size = New System.Drawing.Size(132, 20)
        Me.TextUser.TabIndex = 2
        '
        'TextPassword
        '
        Me.TextPassword.Location = New System.Drawing.Point(209, 90)
        Me.TextPassword.Name = "TextPassword"
        Me.TextPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextPassword.Size = New System.Drawing.Size(132, 20)
        Me.TextPassword.TabIndex = 3
        '
        'BTNOK
        '
        Me.BTNOK.Location = New System.Drawing.Point(265, 165)
        Me.BTNOK.Name = "BTNOK"
        Me.BTNOK.Size = New System.Drawing.Size(75, 32)
        Me.BTNOK.TabIndex = 4
        Me.BTNOK.Text = "ตกลง"
        Me.BTNOK.UseVisualStyleBackColor = True
        '
        'BTNExit
        '
        Me.BTNExit.Location = New System.Drawing.Point(354, 165)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(75, 32)
        Me.BTNExit.TabIndex = 5
        Me.BTNExit.Text = "ออก"
        Me.BTNExit.UseVisualStyleBackColor = True
        '
        'FormLogIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(515, 289)
        Me.Controls.Add(Me.BTNExit)
        Me.Controls.Add(Me.BTNOK)
        Me.Controls.Add(Me.TextPassword)
        Me.Controls.Add(Me.TextUser)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FormLogIN"
        Me.Text = "กรอกชื่อและรหัสผ่านเข้าใช้งานโปรแกรม"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextUser As System.Windows.Forms.TextBox
    Friend WithEvents TextPassword As System.Windows.Forms.TextBox
    Friend WithEvents BTNOK As System.Windows.Forms.Button
    Friend WithEvents BTNExit As System.Windows.Forms.Button
End Class
