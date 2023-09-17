<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Actualizarcomunicados
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Actualizarcomunicados))
        Me.AxAcroPDF3 = New AxAcroPDFLib.AxAcroPDF()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.CheckBox15 = New System.Windows.Forms.CheckBox()
        Me.TextBox29 = New System.Windows.Forms.TextBox()
        Me.LinkLabel15 = New System.Windows.Forms.LinkLabel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.ComboBox6 = New System.Windows.Forms.ComboBox()
        CType(Me.AxAcroPDF3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxAcroPDF3
        '
        Me.AxAcroPDF3.Enabled = True
        Me.AxAcroPDF3.Location = New System.Drawing.Point(6, 60)
        Me.AxAcroPDF3.Name = "AxAcroPDF3"
        Me.AxAcroPDF3.OcxState = CType(resources.GetObject("AxAcroPDF3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxAcroPDF3.Size = New System.Drawing.Size(664, 357)
        Me.AxAcroPDF3.TabIndex = 93
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Green
        Me.Label25.Location = New System.Drawing.Point(157, 9)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(513, 48)
        Me.Label25.TabIndex = 92
        Me.Label25.Text = "En esta pestaña podrás anexar todos los documentos como" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "          permisos, inca" &
    "pacidades, repuestas, entre otros"
        '
        'CheckBox15
        '
        Me.CheckBox15.AutoSize = True
        Me.CheckBox15.Location = New System.Drawing.Point(200, 444)
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox15.TabIndex = 462
        Me.CheckBox15.UseVisualStyleBackColor = True
        '
        'TextBox29
        '
        Me.TextBox29.Location = New System.Drawing.Point(364, 441)
        Me.TextBox29.Name = "TextBox29"
        Me.TextBox29.ReadOnly = True
        Me.TextBox29.Size = New System.Drawing.Size(30, 20)
        Me.TextBox29.TabIndex = 461
        Me.TextBox29.Visible = False
        '
        'LinkLabel15
        '
        Me.LinkLabel15.AutoSize = True
        Me.LinkLabel15.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel15.LinkColor = System.Drawing.Color.Green
        Me.LinkLabel15.Location = New System.Drawing.Point(212, 441)
        Me.LinkLabel15.Name = "LinkLabel15"
        Me.LinkLabel15.Size = New System.Drawing.Size(146, 20)
        Me.LinkLabel15.TabIndex = 2
        Me.LinkLabel15.TabStop = True
        Me.LinkLabel15.Text = "Seleccionar archivo"
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Green
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(583, 435)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(87, 30)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "Cancelar"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Green
        Me.Button6.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.White
        Me.Button6.Location = New System.Drawing.Point(400, 435)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(89, 30)
        Me.Button6.TabIndex = 3
        Me.Button6.Text = "Actualizar"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Green
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(495, 435)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 30)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Eliminar"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Green
        Me.Label42.Location = New System.Drawing.Point(44, 435)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(117, 18)
        Me.Label42.TabIndex = 496
        Me.Label42.Text = "# Comunicado"
        '
        'ComboBox6
        '
        Me.ComboBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox6.FormattingEnabled = True
        Me.ComboBox6.Location = New System.Drawing.Point(66, 456)
        Me.ComboBox6.Name = "ComboBox6"
        Me.ComboBox6.Size = New System.Drawing.Size(73, 21)
        Me.ComboBox6.TabIndex = 1
        '
        'Actualizarcomunicados
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(682, 497)
        Me.Controls.Add(Me.ComboBox6)
        Me.Controls.Add(Me.Label42)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox15)
        Me.Controls.Add(Me.TextBox29)
        Me.Controls.Add(Me.LinkLabel15)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.AxAcroPDF3)
        Me.Controls.Add(Me.Label25)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "Actualizarcomunicados"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Comunicados"
        CType(Me.AxAcroPDF3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents AxAcroPDF3 As AxAcroPDFLib.AxAcroPDF
    Friend WithEvents Label25 As Label
    Friend WithEvents CheckBox15 As CheckBox
    Friend WithEvents TextBox29 As TextBox
    Friend WithEvents LinkLabel15 As LinkLabel
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Label42 As Label
    Friend WithEvents ComboBox6 As ComboBox
End Class
