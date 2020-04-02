<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.rdbtninventarios = New System.Windows.Forms.RadioButton
        Me.rdbtnventas = New System.Windows.Forms.RadioButton
        Me.Button4 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(36, 246)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(160, 245)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 86)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(260, 136)
        Me.TextBox1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Iniciando proceso...."
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(170, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(92, 32)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Iniciar"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'rdbtninventarios
        '
        Me.rdbtninventarios.AutoSize = True
        Me.rdbtninventarios.Location = New System.Drawing.Point(170, 54)
        Me.rdbtninventarios.Name = "rdbtninventarios"
        Me.rdbtninventarios.Size = New System.Drawing.Size(76, 17)
        Me.rdbtninventarios.TabIndex = 6
        Me.rdbtninventarios.Text = "inventarios"
        Me.rdbtninventarios.UseVisualStyleBackColor = True
        Me.rdbtninventarios.Visible = False
        '
        'rdbtnventas
        '
        Me.rdbtnventas.AutoSize = True
        Me.rdbtnventas.Checked = True
        Me.rdbtnventas.Location = New System.Drawing.Point(29, 54)
        Me.rdbtnventas.Name = "rdbtnventas"
        Me.rdbtnventas.Size = New System.Drawing.Size(58, 17)
        Me.rdbtnventas.TabIndex = 7
        Me.rdbtnventas.TabStop = True
        Me.rdbtnventas.Text = "Ventas"
        Me.rdbtnventas.UseVisualStyleBackColor = True
        Me.rdbtnventas.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(288, 54)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 26)
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "Prueba PH"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(380, 86)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.rdbtnventas)
        Me.Controls.Add(Me.rdbtninventarios)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Descarga de Ventas PH"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents rdbtninventarios As System.Windows.Forms.RadioButton
    Friend WithEvents rdbtnventas As System.Windows.Forms.RadioButton
    Friend WithEvents Button4 As System.Windows.Forms.Button
End Class
