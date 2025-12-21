<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Settings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.LblOVLprec = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblBaseDivPrec64bit = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown4 = New System.Windows.Forms.NumericUpDown()
        Me.NumericUpDown3 = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.OkBtn = New System.Windows.Forms.Button()
        Me.CancelBtn = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.txOVLprec = New System.Windows.Forms.NumericUpDown()
        Me.NumericUpDown5 = New System.Windows.Forms.NumericUpDown()
        Me.NumericUpDown6 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txOVLprec, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblOVLprec
        '
        Me.LblOVLprec.AutoSize = True
        Me.LblOVLprec.Location = New System.Drawing.Point(110, 173)
        Me.LblOVLprec.Name = "LblOVLprec"
        Me.LblOVLprec.Size = New System.Drawing.Size(265, 32)
        Me.LblOVLprec.TabIndex = 0
        Me.LblOVLprec.Text = "Round all results to:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 115)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(249, 32)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Extra precision #2:"
        '
        'lblBaseDivPrec64bit
        '
        Me.lblBaseDivPrec64bit.AutoSize = True
        Me.lblBaseDivPrec64bit.Location = New System.Drawing.Point(25, 84)
        Me.lblBaseDivPrec64bit.Name = "lblBaseDivPrec64bit"
        Me.lblBaseDivPrec64bit.Size = New System.Drawing.Size(275, 32)
        Me.lblBaseDivPrec64bit.TabIndex = 2
        Me.lblBaseDivPrec64bit.Text = "By an 8-byte integer:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(249, 32)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Extra precision #1:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.NumericUpDown2)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.lblBaseDivPrec64bit)
        Me.GroupBox1.Location = New System.Drawing.Point(73, 315)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(660, 410)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Division precision"
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown2.Location = New System.Drawing.Point(344, 79)
        Me.NumericUpDown2.Maximum = New Decimal(New Integer() {200000, 0, 0, 0})
        Me.NumericUpDown2.Minimum = New Decimal(New Integer() {200000, 0, 0, -2147483648})
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(135, 41)
        Me.NumericUpDown2.TabIndex = 15
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.NumericUpDown4)
        Me.GroupBox2.Controls.Add(Me.NumericUpDown3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Location = New System.Drawing.Point(43, 149)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(562, 188)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "For big numbers"
        '
        'NumericUpDown4
        '
        Me.NumericUpDown4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown4.Location = New System.Drawing.Point(301, 106)
        Me.NumericUpDown4.Maximum = New Decimal(New Integer() {200000, 0, 0, 0})
        Me.NumericUpDown4.Minimum = New Decimal(New Integer() {200000, 0, 0, -2147483648})
        Me.NumericUpDown4.Name = "NumericUpDown4"
        Me.NumericUpDown4.Size = New System.Drawing.Size(135, 41)
        Me.NumericUpDown4.TabIndex = 17
        '
        'NumericUpDown3
        '
        Me.NumericUpDown3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown3.Location = New System.Drawing.Point(301, 51)
        Me.NumericUpDown3.Maximum = New Decimal(New Integer() {200000, 0, 0, 0})
        Me.NumericUpDown3.Minimum = New Decimal(New Integer() {200000, 0, 0, -2147483648})
        Me.NumericUpDown3.Name = "NumericUpDown3"
        Me.NumericUpDown3.Size = New System.Drawing.Size(135, 41)
        Me.NumericUpDown3.TabIndex = 16
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(89, 981)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(463, 32)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Number of decimal digits displayed:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(65, 814)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(494, 32)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Precision for conversion from decimal:"
        '
        'OkBtn
        '
        Me.OkBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OkBtn.Location = New System.Drawing.Point(562, 1098)
        Me.OkBtn.Name = "OkBtn"
        Me.OkBtn.Size = New System.Drawing.Size(171, 93)
        Me.OkBtn.TabIndex = 7
        Me.OkBtn.Text = "OK"
        Me.OkBtn.UseVisualStyleBackColor = True
        '
        'CancelBtn
        '
        Me.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBtn.Location = New System.Drawing.Point(349, 1098)
        Me.CancelBtn.Name = "CancelBtn"
        Me.CancelBtn.Size = New System.Drawing.Size(171, 93)
        Me.CancelBtn.TabIndex = 8
        Me.CancelBtn.Text = "Cancel"
        Me.CancelBtn.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(83, 88)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(328, 36)
        Me.CheckBox1.TabIndex = 9
        Me.CheckBox1.Text = "Limit overall precision"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'txOVLprec
        '
        Me.txOVLprec.Enabled = False
        Me.txOVLprec.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txOVLprec.Location = New System.Drawing.Point(400, 168)
        Me.txOVLprec.Maximum = New Decimal(New Integer() {200000, 0, 0, 0})
        Me.txOVLprec.Minimum = New Decimal(New Integer() {200000, 0, 0, -2147483648})
        Me.txOVLprec.Name = "txOVLprec"
        Me.txOVLprec.Size = New System.Drawing.Size(135, 41)
        Me.txOVLprec.TabIndex = 14
        '
        'NumericUpDown5
        '
        Me.NumericUpDown5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown5.Location = New System.Drawing.Point(575, 809)
        Me.NumericUpDown5.Maximum = New Decimal(New Integer() {200000, 0, 0, 0})
        Me.NumericUpDown5.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown5.Name = "NumericUpDown5"
        Me.NumericUpDown5.Size = New System.Drawing.Size(135, 41)
        Me.NumericUpDown5.TabIndex = 18
        Me.NumericUpDown5.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'NumericUpDown6
        '
        Me.NumericUpDown6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown6.Increment = New Decimal(New Integer() {18, 0, 0, 0})
        Me.NumericUpDown6.Location = New System.Drawing.Point(575, 976)
        Me.NumericUpDown6.Maximum = New Decimal(New Integer() {18000000, 0, 0, 0})
        Me.NumericUpDown6.Minimum = New Decimal(New Integer() {18, 0, 0, 0})
        Me.NumericUpDown6.Name = "NumericUpDown6"
        Me.NumericUpDown6.Size = New System.Drawing.Size(135, 41)
        Me.NumericUpDown6.TabIndex = 19
        Me.NumericUpDown6.Value = New Decimal(New Integer() {18, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 920)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(562, 32)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "the following one is in digits (rounded to 18)"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 874)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(745, 32)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "All the above values are expressed in QWORDS (8 bytes),"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(49, 1115)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(253, 76)
        Me.Button1.TabIndex = 22
        Me.Button1.Text = "Reset Defaults"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(16.0!, 31.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(800, 1203)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.NumericUpDown6)
        Me.Controls.Add(Me.NumericUpDown5)
        Me.Controls.Add(Me.txOVLprec)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CancelBtn)
        Me.Controls.Add(Me.OkBtn)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.LblOVLprec)
        Me.Name = "Settings"
        Me.Text = "Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txOVLprec, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblOVLprec As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents lblBaseDivPrec64bit As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents OkBtn As Button
    Friend WithEvents CancelBtn As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents NumericUpDown2 As NumericUpDown
    Friend WithEvents NumericUpDown5 As NumericUpDown
    Friend WithEvents NumericUpDown4 As NumericUpDown
    Friend WithEvents NumericUpDown3 As NumericUpDown
    Friend WithEvents txOVLprec As NumericUpDown
    Friend WithEvents NumericUpDown6 As NumericUpDown
    Friend WithEvents Label1 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents ToolTip1 As ToolTip
End Class
