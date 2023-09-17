<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class menum
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(menum))
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Salir = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ajustesPre = New System.Windows.Forms.Button()
        Me.prestamos = New System.Windows.Forms.Button()
        Me.bNomina = New System.Windows.Forms.Button()
        Me.hojaDeVida = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(199, 193)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 30)
        Me.Label3.TabIndex = 39
        Me.Label3.Text = "      Valores " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "predeterminados"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(222, 308)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 15)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "Préstamos"
        '
        'Salir
        '
        Me.Salir.BackColor = System.Drawing.Color.Green
        Me.Salir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Salir.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Salir.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Salir.ForeColor = System.Drawing.Color.White
        Me.Salir.Location = New System.Drawing.Point(76, 353)
        Me.Salir.Name = "Salir"
        Me.Salir.Size = New System.Drawing.Size(212, 30)
        Me.Salir.TabIndex = 35
        Me.Salir.Text = "Salir"
        Me.Salir.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(53, 310)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 15)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Nómina"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(42, 191)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 15)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Hoja de vida"
        '
        'ajustesPre
        '
        Me.ajustesPre.BackgroundImage = CType(resources.GetObject("ajustesPre.BackgroundImage"), System.Drawing.Image)
        Me.ajustesPre.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ajustesPre.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ajustesPre.Location = New System.Drawing.Point(202, 115)
        Me.ajustesPre.Name = "ajustesPre"
        Me.ajustesPre.Size = New System.Drawing.Size(107, 75)
        Me.ajustesPre.TabIndex = 38
        Me.ajustesPre.UseVisualStyleBackColor = True
        '
        'prestamos
        '
        Me.prestamos.BackgroundImage = CType(resources.GetObject("prestamos.BackgroundImage"), System.Drawing.Image)
        Me.prestamos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.prestamos.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.prestamos.Location = New System.Drawing.Point(206, 232)
        Me.prestamos.Name = "prestamos"
        Me.prestamos.Size = New System.Drawing.Size(109, 75)
        Me.prestamos.TabIndex = 36
        Me.prestamos.UseVisualStyleBackColor = True
        '
        'bNomina
        '
        Me.bNomina.BackgroundImage = CType(resources.GetObject("bNomina.BackgroundImage"), System.Drawing.Image)
        Me.bNomina.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.bNomina.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.bNomina.Location = New System.Drawing.Point(33, 232)
        Me.bNomina.Name = "bNomina"
        Me.bNomina.Size = New System.Drawing.Size(101, 75)
        Me.bNomina.TabIndex = 29
        Me.bNomina.UseVisualStyleBackColor = True
        '
        'hojaDeVida
        '
        Me.hojaDeVida.BackgroundImage = CType(resources.GetObject("hojaDeVida.BackgroundImage"), System.Drawing.Image)
        Me.hojaDeVida.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.hojaDeVida.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.hojaDeVida.Location = New System.Drawing.Point(33, 115)
        Me.hojaDeVida.Name = "hojaDeVida"
        Me.hojaDeVida.Size = New System.Drawing.Size(101, 75)
        Me.hojaDeVida.TabIndex = 27
        Me.hojaDeVida.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Location = New System.Drawing.Point(7, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(330, 101)
        Me.Panel1.TabIndex = 26
        '
        'menum
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 395)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ajustesPre)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.prestamos)
        Me.Controls.Add(Me.Salir)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.bNomina)
        Me.Controls.Add(Me.hojaDeVida)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "menum"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Menú"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label3 As Label
    Friend WithEvents ajustesPre As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents prestamos As Button
    Friend WithEvents Salir As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents bNomina As Button
    Friend WithEvents hojaDeVida As Button
    Friend WithEvents Panel1 As Panel
End Class
