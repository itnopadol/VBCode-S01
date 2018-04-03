<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FormLogIn
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
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.CMBProfit = New System.Windows.Forms.ComboBox
        Me.TBPassword = New System.Windows.Forms.TextBox
        Me.TBUserID = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.BTNOK = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.White
        Me.Panel4.Controls.Add(Me.CMBProfit)
        Me.Panel4.Controls.Add(Me.TBPassword)
        Me.Panel4.Controls.Add(Me.TBUserID)
        Me.Panel4.Controls.Add(Me.Label3)
        Me.Panel4.Controls.Add(Me.Label2)
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Controls.Add(Me.BTNOK)
        Me.Panel4.Controls.Add(Me.BTNExit)
        Me.Panel4.Location = New System.Drawing.Point(0, 69)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(325, 231)
        '
        'CMBProfit
        '
        Me.CMBProfit.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.CMBProfit.Items.Add("S01")
        Me.CMBProfit.Items.Add("S02")
        Me.CMBProfit.Location = New System.Drawing.Point(197, 28)
        Me.CMBProfit.Name = "CMBProfit"
        Me.CMBProfit.Size = New System.Drawing.Size(103, 26)
        Me.CMBProfit.TabIndex = 0
        '
        'TBPassword
        '
        Me.TBPassword.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TBPassword.Location = New System.Drawing.Point(197, 97)
        Me.TBPassword.Name = "TBPassword"
        Me.TBPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TBPassword.Size = New System.Drawing.Size(102, 26)
        Me.TBPassword.TabIndex = 2
        '
        'TBUserID
        '
        Me.TBUserID.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TBUserID.Location = New System.Drawing.Point(197, 62)
        Me.TBUserID.Name = "TBUserID"
        Me.TBUserID.Size = New System.Drawing.Size(102, 26)
        Me.TBUserID.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(120, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 21)
        Me.Label3.Text = "รหัสผ่าน :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(95, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 21)
        Me.Label2.Text = "รหัสผู้ใช้งาน :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(70, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 21)
        Me.Label1.Text = "บริษัทเข้าใช้งาน :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'BTNOK
        '
        Me.BTNOK.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.BTNOK.Location = New System.Drawing.Point(197, 137)
        Me.BTNOK.Name = "BTNOK"
        Me.BTNOK.Size = New System.Drawing.Size(102, 33)
        Me.BTNOK.TabIndex = 3
        Me.BTNOK.Text = "OK-ตกลง"
        '
        'BTNExit
        '
        Me.BTNExit.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.BTNExit.Location = New System.Drawing.Point(197, 176)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(102, 33)
        Me.BTNExit.TabIndex = 4
        Me.BTNExit.Text = "ESC-ออก"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(2, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(320, 21)
        Me.Label4.Text = "กรอกรหัสพนักงานและรหัสผ่านของ BC"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(1, 10)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(197, 21)
        Me.Label5.Text = "บริษัท นพดลพานิช จำกัด"
        '
        'FormLogIn
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.MediumBlue
        Me.ClientSize = New System.Drawing.Size(325, 300)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormLogIn"
        Me.Text = "FormLogIn"
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents BTNOK As System.Windows.Forms.Button
    Friend WithEvents CMBProfit As System.Windows.Forms.ComboBox
    Friend WithEvents TBPassword As System.Windows.Forms.TextBox
    Friend WithEvents TBUserID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
