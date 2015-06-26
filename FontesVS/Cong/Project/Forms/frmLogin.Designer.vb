<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
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
        Me.cmdEntrar = New System.Windows.Forms.Button()
        Me.txtEmpresa = New System.Windows.Forms.TextBox()
        Me.txtUsuario = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblEmp = New System.Windows.Forms.Label()
        Me.lblUsuario = New System.Windows.Forms.Label()
        Me.lblSenha = New System.Windows.Forms.Label()
        Me.txtSenha = New System.Windows.Forms.MaskedTextBox()
        Me.SuspendLayout()
        '
        'cmdEntrar
        '
        Me.cmdEntrar.Location = New System.Drawing.Point(12, 108)
        Me.cmdEntrar.Name = "cmdEntrar"
        Me.cmdEntrar.Size = New System.Drawing.Size(104, 23)
        Me.cmdEntrar.TabIndex = 4
        Me.cmdEntrar.Text = "Entrar"
        Me.cmdEntrar.UseVisualStyleBackColor = True
        '
        'txtEmpresa
        '
        Me.txtEmpresa.Location = New System.Drawing.Point(119, 13)
        Me.txtEmpresa.MaxLength = 2
        Me.txtEmpresa.Name = "txtEmpresa"
        Me.txtEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtEmpresa.Size = New System.Drawing.Size(18, 20)
        Me.txtEmpresa.TabIndex = 1
        '
        'txtUsuario
        '
        Me.txtUsuario.Location = New System.Drawing.Point(119, 39)
        Me.txtUsuario.MaxLength = 40
        Me.txtUsuario.Name = "txtUsuario"
        Me.txtUsuario.Size = New System.Drawing.Size(100, 20)
        Me.txtUsuario.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(122, 108)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Sair"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lblEmp
        '
        Me.lblEmp.AutoSize = True
        Me.lblEmp.Location = New System.Drawing.Point(16, 16)
        Me.lblEmp.Name = "lblEmp"
        Me.lblEmp.Size = New System.Drawing.Size(51, 13)
        Me.lblEmp.TabIndex = 5
        Me.lblEmp.Text = "Empresa:"
        '
        'lblUsuario
        '
        Me.lblUsuario.AutoSize = True
        Me.lblUsuario.Location = New System.Drawing.Point(16, 42)
        Me.lblUsuario.Name = "lblUsuario"
        Me.lblUsuario.Size = New System.Drawing.Size(46, 13)
        Me.lblUsuario.TabIndex = 6
        Me.lblUsuario.Text = "Usuário:"
        '
        'lblSenha
        '
        Me.lblSenha.AutoSize = True
        Me.lblSenha.Location = New System.Drawing.Point(16, 68)
        Me.lblSenha.Name = "lblSenha"
        Me.lblSenha.Size = New System.Drawing.Size(41, 13)
        Me.lblSenha.TabIndex = 7
        Me.lblSenha.Text = "Senha:"
        '
        'txtSenha
        '
        Me.txtSenha.Location = New System.Drawing.Point(119, 65)
        Me.txtSenha.Name = "txtSenha"
        Me.txtSenha.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSenha.Size = New System.Drawing.Size(100, 20)
        Me.txtSenha.TabIndex = 3
        '
        'frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(233, 141)
        Me.Controls.Add(Me.txtSenha)
        Me.Controls.Add(Me.lblSenha)
        Me.Controls.Add(Me.lblUsuario)
        Me.Controls.Add(Me.lblEmp)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtUsuario)
        Me.Controls.Add(Me.txtEmpresa)
        Me.Controls.Add(Me.cmdEntrar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdEntrar As System.Windows.Forms.Button
    Friend WithEvents txtEmpresa As System.Windows.Forms.TextBox
    Friend WithEvents txtUsuario As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblEmp As System.Windows.Forms.Label
    Friend WithEvents lblUsuario As System.Windows.Forms.Label
    Friend WithEvents lblSenha As System.Windows.Forms.Label
    Friend WithEvents txtSenha As System.Windows.Forms.MaskedTextBox

End Class
