<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class popupForm
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
        Me.rtb01 = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'rtb01
        '
        Me.rtb01.Location = New System.Drawing.Point(12, 12)
        Me.rtb01.Name = "rtb01"
        Me.rtb01.Size = New System.Drawing.Size(115, 123)
        Me.rtb01.TabIndex = 0
        Me.rtb01.Text = ""
        '
        'popupForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(161, 147)
        Me.Controls.Add(Me.rtb01)
        Me.Name = "popupForm"
        Me.Text = "popupForm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents rtb01 As RichTextBox
End Class
