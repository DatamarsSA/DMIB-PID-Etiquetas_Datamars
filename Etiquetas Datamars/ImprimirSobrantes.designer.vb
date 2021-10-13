<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImprimirSobrantes
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.botonSi = New System.Windows.Forms.Button()
        Me.botonNo = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(155, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(148, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Pedido con Sobrantes"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(80, 114)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(296, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "¿Desea imprimir para un total de un 2% más?"
        '
        'botonSi
        '
        Me.botonSi.Location = New System.Drawing.Point(287, 204)
        Me.botonSi.Name = "botonSi"
        Me.botonSi.Size = New System.Drawing.Size(75, 23)
        Me.botonSi.TabIndex = 2
        Me.botonSi.Text = "SI"
        Me.botonSi.UseVisualStyleBackColor = True
        '
        'botonNo
        '
        Me.botonNo.Location = New System.Drawing.Point(101, 204)
        Me.botonNo.Name = "botonNo"
        Me.botonNo.Size = New System.Drawing.Size(75, 23)
        Me.botonNo.TabIndex = 3
        Me.botonNo.Text = "NO"
        Me.botonNo.UseVisualStyleBackColor = True
        '
        'ImprimirSobrantes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(464, 311)
        Me.Controls.Add(Me.botonNo)
        Me.Controls.Add(Me.botonSi)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ImprimirSobrantes"
        Me.Text = "ImprimirSobrantes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents botonSi As Button
    Friend WithEvents botonNo As Button
End Class
