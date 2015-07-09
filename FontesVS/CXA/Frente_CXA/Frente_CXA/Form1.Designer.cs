namespace Frente_CXA
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menu = new System.Windows.Forms.MenuStrip();
            this.tabelasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.operaçõesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.consultasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.administraçãoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.caixaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.abrirCaixaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fecharCaixaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.operadoresToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sobreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menu.SuspendLayout();
            this.SuspendLayout();
            // 
            // menu
            // 
            this.menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tabelasToolStripMenuItem,
            this.operaçõesToolStripMenuItem,
            this.consultasToolStripMenuItem,
            this.administraçãoToolStripMenuItem});
            this.menu.Location = new System.Drawing.Point(0, 0);
            this.menu.Name = "menu";
            this.menu.Size = new System.Drawing.Size(539, 24);
            this.menu.TabIndex = 0;
            this.menu.Text = "menuStrip1";
            // 
            // tabelasToolStripMenuItem
            // 
            this.tabelasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.operadoresToolStripMenuItem});
            this.tabelasToolStripMenuItem.Name = "tabelasToolStripMenuItem";
            this.tabelasToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.tabelasToolStripMenuItem.Text = "Tabelas";
            // 
            // operaçõesToolStripMenuItem
            // 
            this.operaçõesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.caixaToolStripMenuItem,
            this.abrirCaixaToolStripMenuItem,
            this.fecharCaixaToolStripMenuItem});
            this.operaçõesToolStripMenuItem.Name = "operaçõesToolStripMenuItem";
            this.operaçõesToolStripMenuItem.Size = new System.Drawing.Size(75, 20);
            this.operaçõesToolStripMenuItem.Text = "Operações";
            // 
            // consultasToolStripMenuItem
            // 
            this.consultasToolStripMenuItem.Name = "consultasToolStripMenuItem";
            this.consultasToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
            this.consultasToolStripMenuItem.Text = "Consultas";
            // 
            // administraçãoToolStripMenuItem
            // 
            this.administraçãoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.sobreToolStripMenuItem});
            this.administraçãoToolStripMenuItem.Name = "administraçãoToolStripMenuItem";
            this.administraçãoToolStripMenuItem.Size = new System.Drawing.Size(96, 20);
            this.administraçãoToolStripMenuItem.Text = "Administração";
            // 
            // caixaToolStripMenuItem
            // 
            this.caixaToolStripMenuItem.Name = "caixaToolStripMenuItem";
            this.caixaToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.caixaToolStripMenuItem.Text = "Caixa";
            // 
            // abrirCaixaToolStripMenuItem
            // 
            this.abrirCaixaToolStripMenuItem.Name = "abrirCaixaToolStripMenuItem";
            this.abrirCaixaToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.abrirCaixaToolStripMenuItem.Text = "Abrir Caixa";
            // 
            // fecharCaixaToolStripMenuItem
            // 
            this.fecharCaixaToolStripMenuItem.Name = "fecharCaixaToolStripMenuItem";
            this.fecharCaixaToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.fecharCaixaToolStripMenuItem.Text = "Fechar Caixa";
            // 
            // operadoresToolStripMenuItem
            // 
            this.operadoresToolStripMenuItem.Name = "operadoresToolStripMenuItem";
            this.operadoresToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.operadoresToolStripMenuItem.Text = "Operadores";
            // 
            // sobreToolStripMenuItem
            // 
            this.sobreToolStripMenuItem.Name = "sobreToolStripMenuItem";
            this.sobreToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.sobreToolStripMenuItem.Text = "Sobre";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 399);
            this.Controls.Add(this.menu);
            this.MainMenuStrip = this.menu;
            this.Name = "Form1";
            this.Text = "Infinity - Frente de Caixa";
            this.menu.ResumeLayout(false);
            this.menu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menu;
        private System.Windows.Forms.ToolStripMenuItem tabelasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem operadoresToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem operaçõesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem caixaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem abrirCaixaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fecharCaixaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem consultasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem administraçãoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sobreToolStripMenuItem;
    }
}

