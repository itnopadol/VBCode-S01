<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgPSVdocSearch
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
        Me.LVps = New System.Windows.Forms.ListView
        Me.hdPsDocno = New System.Windows.Forms.ColumnHeader
        Me.hdPsDocdate = New System.Windows.Forms.ColumnHeader
        Me.hdOwnerCode = New System.Windows.Forms.ColumnHeader
        Me.hdUserId = New System.Windows.Forms.ColumnHeader
        Me.hdOwnername = New System.Windows.Forms.ColumnHeader
        Me.GB02 = New System.Windows.Forms.GroupBox
        Me.btnOKps = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtPSDoc = New System.Windows.Forms.TextBox
        Me.btnImport = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.cbxAll = New System.Windows.Forms.CheckBox
        Me.GB02.SuspendLayout()
        Me.SuspendLayout()
        '
        'LVps
        '
        Me.LVps.CheckBoxes = True
        Me.LVps.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.hdPsDocno, Me.hdPsDocdate, Me.hdOwnerCode, Me.hdUserId, Me.hdOwnername})
        Me.LVps.FullRowSelect = True
        Me.LVps.GridLines = True
        Me.LVps.Location = New System.Drawing.Point(12, 81)
        Me.LVps.Name = "LVps"
        Me.LVps.Size = New System.Drawing.Size(833, 432)
        Me.LVps.TabIndex = 1
        Me.LVps.UseCompatibleStateImageBehavior = False
        Me.LVps.View = System.Windows.Forms.View.Details
        '
        'hdPsDocno
        '
        Me.hdPsDocno.Text = "เลขที่เอกสารโครงสร้างราคา"
        Me.hdPsDocno.Width = 179
        '
        'hdPsDocdate
        '
        Me.hdPsDocdate.Text = "วันที่เอกสาร"
        Me.hdPsDocdate.Width = 140
        '
        'hdOwnerCode
        '
        Me.hdOwnerCode.Text = "รหัสผู้กำหนด"
        Me.hdOwnerCode.Width = 134
        '
        'hdUserId
        '
        Me.hdUserId.Text = "UserId"
        '
        'hdOwnername
        '
        Me.hdOwnername.Text = "ผู้กำหนด"
        Me.hdOwnername.Width = 277
        '
        'GB02
        '
        Me.GB02.BackColor = System.Drawing.SystemColors.Desktop
        Me.GB02.Controls.Add(Me.btnOKps)
        Me.GB02.Controls.Add(Me.Label10)
        Me.GB02.Controls.Add(Me.txtPSDoc)
        Me.GB02.ForeColor = System.Drawing.Color.White
        Me.GB02.Location = New System.Drawing.Point(12, 12)
        Me.GB02.Name = "GB02"
        Me.GB02.Size = New System.Drawing.Size(740, 62)
        Me.GB02.TabIndex = 18
        Me.GB02.TabStop = False
        Me.GB02.Text = "กรุณากรอกคำค้นหาในช่องคำค้นหา"
        '
        'btnOKps
        '
        Me.btnOKps.ForeColor = System.Drawing.Color.Black
        Me.btnOKps.Location = New System.Drawing.Point(528, 24)
        Me.btnOKps.Name = "btnOKps"
        Me.btnOKps.Size = New System.Drawing.Size(75, 28)
        Me.btnOKps.TabIndex = 2
        Me.btnOKps.Text = "ตกลง"
        Me.btnOKps.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label10.Location = New System.Drawing.Point(69, 30)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(242, 16)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "คำค้นหา หรือ เลขที่เอกสารโครงสร้างราคา :"
        '
        'txtPSDoc
        '
        Me.txtPSDoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtPSDoc.Location = New System.Drawing.Point(313, 26)
        Me.txtPSDoc.Name = "txtPSDoc"
        Me.txtPSDoc.Size = New System.Drawing.Size(214, 26)
        Me.txtPSDoc.TabIndex = 0
        '
        'btnImport
        '
        'Me.btnImport.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_25
        Me.btnImport.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnImport.Location = New System.Drawing.Point(658, 519)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(93, 53)
        Me.btnImport.TabIndex = 19
        Me.btnImport.Text = "นำเข้าข้อมูล"
        Me.btnImport.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'Button3
        '
        'Me.Button3.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_73
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button3.Location = New System.Drawing.Point(750, 519)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(95, 53)
        Me.Button3.TabIndex = 20
        Me.Button3.Text = "ยกเลิก"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button3.UseVisualStyleBackColor = True
        '
        'cbxAll
        '
        Me.cbxAll.AutoSize = True
        Me.cbxAll.ForeColor = System.Drawing.Color.White
        Me.cbxAll.Location = New System.Drawing.Point(12, 519)
        Me.cbxAll.Name = "cbxAll"
        Me.cbxAll.Size = New System.Drawing.Size(83, 17)
        Me.cbxAll.TabIndex = 21
        Me.cbxAll.Text = "เลือกทั้งหมด"
        Me.cbxAll.UseVisualStyleBackColor = True
        '
        'dlgPSVdocSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(863, 576)
        Me.Controls.Add(Me.cbxAll)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.GB02)
        Me.Controls.Add(Me.LVps)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgPSVdocSearch"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ค้นหาเอกสารโครงสร้างราคา"
        Me.GB02.ResumeLayout(False)
        Me.GB02.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LVps As System.Windows.Forms.ListView
    Friend WithEvents GB02 As System.Windows.Forms.GroupBox
    Friend WithEvents btnOKps As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtPSDoc As System.Windows.Forms.TextBox
    Friend WithEvents hdPsDocno As System.Windows.Forms.ColumnHeader
    Friend WithEvents hdPsDocdate As System.Windows.Forms.ColumnHeader
    Friend WithEvents hdOwnerCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents hdOwnername As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents cbxAll As System.Windows.Forms.CheckBox
    Friend WithEvents hdUserId As System.Windows.Forms.ColumnHeader

End Class
