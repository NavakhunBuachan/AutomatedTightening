<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class debugForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.tbSend = New System.Windows.Forms.TextBox()
        Me.rtbRecv = New System.Windows.Forms.RichTextBox()
        Me.btnInit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(250, 105)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(75, 23)
        Me.btnSend.TabIndex = 0
        Me.btnSend.Text = "Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'tbSend
        '
        Me.tbSend.Location = New System.Drawing.Point(46, 107)
        Me.tbSend.Name = "tbSend"
        Me.tbSend.Size = New System.Drawing.Size(171, 20)
        Me.tbSend.TabIndex = 1
        '
        'rtbRecv
        '
        Me.rtbRecv.Location = New System.Drawing.Point(46, 151)
        Me.rtbRecv.Name = "rtbRecv"
        Me.rtbRecv.Size = New System.Drawing.Size(171, 96)
        Me.rtbRecv.TabIndex = 2
        Me.rtbRecv.Text = ""
        '
        'btnInit
        '
        Me.btnInit.Location = New System.Drawing.Point(46, 51)
        Me.btnInit.Name = "btnInit"
        Me.btnInit.Size = New System.Drawing.Size(75, 23)
        Me.btnInit.TabIndex = 3
        Me.btnInit.Text = "RS232 Initial"
        Me.btnInit.UseVisualStyleBackColor = True
        '
        'debugForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnInit)
        Me.Controls.Add(Me.rtbRecv)
        Me.Controls.Add(Me.tbSend)
        Me.Controls.Add(Me.btnSend)
        Me.Name = "debugForm"
        Me.Text = "debugForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnSend As Button
    Friend WithEvents tbSend As TextBox
    Friend WithEvents rtbRecv As RichTextBox
    Friend WithEvents btnInit As Button
End Class
