<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class mainForm
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
        Me.tbSn = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnMESCheck = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lbOutput = New System.Windows.Forms.Label()
        Me.btnFin = New System.Windows.Forms.Button()
        Me.lbOutput2 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbEn = New System.Windows.Forms.TextBox()
        Me.btnInit = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbSn
        '
        Me.tbSn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbSn.Location = New System.Drawing.Point(25, 53)
        Me.tbSn.Name = "tbSn"
        Me.tbSn.Size = New System.Drawing.Size(228, 26)
        Me.tbSn.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Serial Number :"
        '
        'btnMESCheck
        '
        Me.btnMESCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMESCheck.Location = New System.Drawing.Point(25, 214)
        Me.btnMESCheck.Name = "btnMESCheck"
        Me.btnMESCheck.Size = New System.Drawing.Size(228, 56)
        Me.btnMESCheck.TabIndex = 2
        Me.btnMESCheck.Text = "MES Check"
        Me.btnMESCheck.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(25, 310)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(228, 56)
        Me.btnStart.TabIndex = 3
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(381, 53)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1396, 784)
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.TabStop = False
        '
        'lbOutput
        '
        Me.lbOutput.AutoSize = True
        Me.lbOutput.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbOutput.Location = New System.Drawing.Point(432, 869)
        Me.lbOutput.Name = "lbOutput"
        Me.lbOutput.Size = New System.Drawing.Size(699, 37)
        Me.lbOutput.TabIndex = 11
        Me.lbOutput.Text = "OutPutLabel........................................................"
        '
        'btnFin
        '
        Me.btnFin.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFin.Location = New System.Drawing.Point(25, 635)
        Me.btnFin.Name = "btnFin"
        Me.btnFin.Size = New System.Drawing.Size(228, 56)
        Me.btnFin.TabIndex = 4
        Me.btnFin.Text = "Finish"
        Me.btnFin.UseVisualStyleBackColor = True
        '
        'lbOutput2
        '
        Me.lbOutput2.AutoSize = True
        Me.lbOutput2.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbOutput2.Location = New System.Drawing.Point(432, 918)
        Me.lbOutput2.Name = "lbOutput2"
        Me.lbOutput2.Size = New System.Drawing.Size(699, 37)
        Me.lbOutput2.TabIndex = 12
        Me.lbOutput2.Text = "OutPutLabel........................................................"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 98)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 20)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Operator ID :"
        '
        'tbEn
        '
        Me.tbEn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbEn.Location = New System.Drawing.Point(25, 131)
        Me.tbEn.Name = "tbEn"
        Me.tbEn.Size = New System.Drawing.Size(228, 26)
        Me.tbEn.TabIndex = 1
        '
        'btnInit
        '
        Me.btnInit.Location = New System.Drawing.Point(41, 814)
        Me.btnInit.Name = "btnInit"
        Me.btnInit.Size = New System.Drawing.Size(101, 23)
        Me.btnInit.TabIndex = 13
        Me.btnInit.Text = "CheckStatus"
        Me.btnInit.UseVisualStyleBackColor = True
        '
        'mainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1869, 1061)
        Me.Controls.Add(Me.btnInit)
        Me.Controls.Add(Me.tbEn)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbOutput2)
        Me.Controls.Add(Me.btnFin)
        Me.Controls.Add(Me.lbOutput)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnMESCheck)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbSn)
        Me.Name = "mainForm"
        Me.Text = "AutomatedScrew"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbSn As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnMESCheck As Button
    Friend WithEvents btnStart As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents lbOutput As Label
    Friend WithEvents btnFin As Button
    Friend WithEvents lbOutput2 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents tbEn As TextBox
    Friend WithEvents btnInit As Button
End Class
