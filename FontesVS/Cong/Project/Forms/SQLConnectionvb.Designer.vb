<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSQLConnection
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
        Me.txtInstancia = New System.Windows.Forms.TextBox()
        Me.txtBanco = New System.Windows.Forms.TextBox()
        Me.txtUsuario = New System.Windows.Forms.TextBox()
        Me.txtSenha = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnAvancar = New System.Windows.Forms.Button()
        Me.lblInstancia = New System.Windows.Forms.Label()
        Me.lblBanco = New System.Windows.Forms.Label()
        Me.lblUsuario = New System.Windows.Forms.Label()
        Me.lblSenha = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtInstancia
        '
        Me.txtInstancia.Location = New System.Drawing.Point(72, 11)
        Me.txtInstancia.Name = "txtInstancia"
        Me.txtInstancia.Size = New System.Drawing.Size(130, 20)
        Me.txtInstancia.TabIndex = 0
        '
        'txtBanco
        '
        Me.txtBanco.Location = New System.Drawing.Point(72, 37)
        Me.txtBanco.Name = "txtBanco"
        Me.txtBanco.Size = New System.Drawing.Size(130, 20)
        Me.txtBanco.TabIndex = 1
        '
        'txtUsuario
        '
        Me.txtUsuario.Location = New System.Drawing.Point(72, 63)
        Me.txtUsuario.Name = "txtUsuario"
        Me.txtUsuario.Size = New System.Drawing.Size(130, 20)
        Me.txtUsuario.TabIndex = 2
        '
        'txtSenha
        '
        Me.txtSenha.Location = New System.Drawing.Point(72, 89)
        Me.txtSenha.Name = "txtSenha"
        Me.txtSenha.Size = New System.Drawing.Size(130, 20)
        Me.txtSenha.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 125)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(92, 25)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "&Sair"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnAvancar
        '
        Me.btnAvancar.Location = New System.Drawing.Point(110, 125)
        Me.btnAvancar.Name = "btnAvancar"
        Me.btnAvancar.Size = New System.Drawing.Size(92, 25)
        Me.btnAvancar.TabIndex = 5
        Me.btnAvancar.Text = "&Avançar"
        Me.btnAvancar.UseVisualStyleBackColor = True
        '
        'lblInstancia
        '
        Me.lblInstancia.AutoSize = True
        Me.lblInstancia.Location = New System.Drawing.Point(9, 14)
        Me.lblInstancia.Name = "lblInstancia"
        Me.lblInstancia.Size = New System.Drawing.Size(53, 13)
        Me.lblInstancia.TabIndex = 6
        Me.lblInstancia.Text = "Instância:"
        '
        'lblBanco
        '
        Me.lblBanco.AutoSize = True
        Me.lblBanco.Location = New System.Drawing.Point(9, 40)
        Me.lblBanco.Name = "lblBanco"
        Me.lblBanco.Size = New System.Drawing.Size(41, 13)
        Me.lblBanco.TabIndex = 7
        Me.lblBanco.Text = "Banco:"
        '
        'lblUsuario
        '
        Me.lblUsuario.AutoSize = True
        Me.lblUsuario.Location = New System.Drawing.Point(9, 66)
        Me.lblUsuario.Name = "lblUsuario"
        Me.lblUsuario.Size = New System.Drawing.Size(46, 13)
        Me.lblUsuario.TabIndex = 8
        Me.lblUsuario.Text = "Usuário:"
        '
        'lblSenha
        '
        Me.lblSenha.AutoSize = True
        Me.lblSenha.Location = New System.Drawing.Point(9, 92)
        Me.lblSenha.Name = "lblSenha"
        Me.lblSenha.Size = New System.Drawing.Size(41, 13)
        Me.lblSenha.TabIndex = 9
        Me.lblSenha.Text = "Senha:"
        '
        'frmSQLConnection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(212, 158)
        Me.Controls.Add(Me.lblSenha)
        Me.Controls.Add(Me.lblUsuario)
        Me.Controls.Add(Me.lblBanco)
        Me.Controls.Add(Me.lblInstancia)
        Me.Controls.Add(Me.btnAvancar)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtSenha)
        Me.Controls.Add(Me.txtUsuario)
        Me.Controls.Add(Me.txtBanco)
        Me.Controls.Add(Me.txtInstancia)
        Me.Name = "frmSQLConnection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SQLConnectionvb"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtInstancia As System.Windows.Forms.TextBox
    Friend WithEvents txtBanco As System.Windows.Forms.TextBox
    Friend WithEvents txtUsuario As System.Windows.Forms.TextBox
    Friend WithEvents txtSenha As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnAvancar As System.Windows.Forms.Button
    Friend WithEvents lblInstancia As System.Windows.Forms.Label
    Friend WithEvents lblBanco As System.Windows.Forms.Label
    Friend WithEvents lblUsuario As System.Windows.Forms.Label
    Friend WithEvents lblSenha As System.Windows.Forms.Label
End Class
