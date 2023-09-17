<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ActualizarCertificadoLaboral
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ActualizarCertificadoLaboral))
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.AxAcroPDF4 = New AxAcroPDFLib.AxAcroPDF()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CheckBox15 = New System.Windows.Forms.CheckBox()
        Me.TextBox29 = New System.Windows.Forms.TextBox()
        Me.LinkLabel15 = New System.Windows.Forms.LinkLabel()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.AxAcroPDF4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.Green
        Me.Button9.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button9.ForeColor = System.Drawing.Color.White
        Me.Button9.Location = New System.Drawing.Point(598, 408)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(87, 30)
        Me.Button9.TabIndex = 5
        Me.Button9.Text = "Cancelar"
        Me.Button9.UseVisualStyleBackColor = False
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.Green
        Me.Button10.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Button10.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10.ForeColor = System.Drawing.Color.White
        Me.Button10.Location = New System.Drawing.Point(411, 408)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(93, 30)
        Me.Button10.TabIndex = 3
        Me.Button10.Text = "Actualizar"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'AxAcroPDF4
        '
        Me.AxAcroPDF4.Enabled = True
        Me.AxAcroPDF4.Location = New System.Drawing.Point(7, 5)
        Me.AxAcroPDF4.Name = "AxAcroPDF4"
        Me.AxAcroPDF4.OcxState = CType(resources.GetObject("AxAcroPDF4.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxAcroPDF4.Size = New System.Drawing.Size(677, 384)
        Me.AxAcroPDF4.TabIndex = 96
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Green
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(510, 408)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 30)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Eliminar"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'CheckBox15
        '
        Me.CheckBox15.AutoSize = True
        Me.CheckBox15.Location = New System.Drawing.Point(197, 417)
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox15.TabIndex = 498
        Me.CheckBox15.UseVisualStyleBackColor = True
        '
        'TextBox29
        '
        Me.TextBox29.Location = New System.Drawing.Point(370, 414)
        Me.TextBox29.Name = "TextBox29"
        Me.TextBox29.ReadOnly = True
        Me.TextBox29.Size = New System.Drawing.Size(35, 20)
        Me.TextBox29.TabIndex = 497
        Me.TextBox29.Visible = False
        '
        'LinkLabel15
        '
        Me.LinkLabel15.AutoSize = True
        Me.LinkLabel15.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel15.LinkColor = System.Drawing.Color.Green
        Me.LinkLabel15.Location = New System.Drawing.Point(218, 414)
        Me.LinkLabel15.Name = "LinkLabel15"
        Me.LinkLabel15.Size = New System.Drawing.Size(146, 20)
        Me.LinkLabel15.TabIndex = 2
        Me.LinkLabel15.TabStop = True
        Me.LinkLabel15.Text = "Seleccionar archivo"
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(59, 417)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(82, 21)
        Me.ComboBox1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Green
        Me.Label1.Location = New System.Drawing.Point(56, 392)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 18)
        Me.Label1.TabIndex = 501
        Me.Label1.Text = "# Certificado"
        '
        'ActualizarCertificadoLaboral
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(696, 460)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CheckBox15)
        Me.Controls.Add(Me.TextBox29)
        Me.Controls.Add(Me.LinkLabel15)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.AxAcroPDF4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "ActualizarCertificadoLaboral"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Certificado laboral"
        CType(Me.AxAcroPDF4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button9 As Button
    Friend WithEvents Button10 As Button
    Friend WithEvents AxAcroPDF4 As AxAcroPDFLib.AxAcroPDF
    Friend WithEvents Button1 As Button
    Friend WithEvents CheckBox15 As CheckBox
    Friend WithEvents TextBox29 As TextBox
    Friend WithEvents LinkLabel15 As LinkLabel
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
End Class
