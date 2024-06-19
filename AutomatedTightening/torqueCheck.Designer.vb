<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class torqueCheck
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnFin = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rtb01
        '
        Me.rtb01.Location = New System.Drawing.Point(45, 420)
        Me.rtb01.Name = "rtb01"
        Me.rtb01.Size = New System.Drawing.Size(175, 98)
        Me.rtb01.TabIndex = 0
        Me.rtb01.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'btnFin
        '
        Me.btnFin.Location = New System.Drawing.Point(521, 504)
        Me.btnFin.Name = "btnFin"
        Me.btnFin.Size = New System.Drawing.Size(75, 23)
        Me.btnFin.TabIndex = 2
        Me.btnFin.Text = "Finish"
        Me.btnFin.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(45, 55)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1071, 342)
        Me.DataGridView1.TabIndex = 3
        '
        'torqueCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1176, 539)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnFin)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.rtb01)
        Me.Name = "torqueCheck"
        Me.Text = "torqueCheck"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rtb01 As RichTextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnFin As Button
    Friend WithEvents DataGridView1 As DataGridView
End Class
