namespace InterfazCi
{
    partial class Base
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
            this.ciCompanyList11 = new controles.CICompanyList1();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPass = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBD = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.buttonversap = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txtpwdSAP = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtUserSAP = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtBDSAP = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.TxtServerSAP = new System.Windows.Forms.TextBox();
            this.ciCompanyList12 = new controles.CICompanyList1();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // ciCompanyList11
            // 
            this.ciCompanyList11.Location = new System.Drawing.Point(5, 20);
            this.ciCompanyList11.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ciCompanyList11.Name = "ciCompanyList11";
            this.ciCompanyList11.Size = new System.Drawing.Size(384, 36);
            this.ciCompanyList11.TabIndex = 3;
            this.ciCompanyList11.Load += new System.EventHandler(this.ciCompanyList11_Load);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(552, 277);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.ciCompanyList11);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.tabPage1.Size = new System.Drawing.Size(544, 251);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Proceso";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.button1);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.txtPass);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.txtUser);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.txtBD);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.txtServer);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.tabPage2.Size = new System.Drawing.Size(544, 251);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Conexion Contabilidad";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(34, 137);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(243, 23);
            this.button1.TabIndex = 19;
            this.button1.Text = "Test y Guardar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(34, 102);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 20);
            this.label4.TabIndex = 18;
            this.label4.Text = "Password";
            // 
            // txtPass
            // 
            this.txtPass.Location = new System.Drawing.Point(139, 102);
            this.txtPass.Name = "txtPass";
            this.txtPass.PasswordChar = '*';
            this.txtPass.Size = new System.Drawing.Size(138, 20);
            this.txtPass.TabIndex = 17;
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(34, 76);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 20);
            this.label3.TabIndex = 16;
            this.label3.Text = "Usuario";
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(139, 76);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(138, 20);
            this.txtUser.TabIndex = 15;
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(34, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 20);
            this.label2.TabIndex = 14;
            this.label2.Text = "Base Datos";
            // 
            // txtBD
            // 
            this.txtBD.Enabled = false;
            this.txtBD.Location = new System.Drawing.Point(139, 48);
            this.txtBD.Name = "txtBD";
            this.txtBD.Size = new System.Drawing.Size(138, 20);
            this.txtBD.TabIndex = 13;
            this.txtBD.Text = "GeneralesSQL";
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(33, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 20);
            this.label1.TabIndex = 12;
            this.label1.Text = "Server";
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(139, 22);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(138, 20);
            this.txtServer.TabIndex = 11;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.buttonversap);
            this.tabPage3.Controls.Add(this.label5);
            this.tabPage3.Controls.Add(this.txtpwdSAP);
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.txtUserSAP);
            this.tabPage3.Controls.Add(this.label7);
            this.tabPage3.Controls.Add(this.txtBDSAP);
            this.tabPage3.Controls.Add(this.label8);
            this.tabPage3.Controls.Add(this.TxtServerSAP);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(544, 251);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "tabPage3";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // buttonversap
            // 
            this.buttonversap.Location = new System.Drawing.Point(25, 131);
            this.buttonversap.Name = "buttonversap";
            this.buttonversap.Size = new System.Drawing.Size(243, 23);
            this.buttonversap.TabIndex = 37;
            this.buttonversap.Text = "Test y Guardar";
            this.buttonversap.UseVisualStyleBackColor = true;
            this.buttonversap.Click += new System.EventHandler(this.button2_Click);
            // 
            // label5
            // 
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(25, 96);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 20);
            this.label5.TabIndex = 36;
            this.label5.Text = "Password";
            // 
            // txtpwdSAP
            // 
            this.txtpwdSAP.Location = new System.Drawing.Point(130, 96);
            this.txtpwdSAP.Name = "txtpwdSAP";
            this.txtpwdSAP.PasswordChar = '*';
            this.txtpwdSAP.Size = new System.Drawing.Size(138, 20);
            this.txtpwdSAP.TabIndex = 35;
            // 
            // label6
            // 
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(25, 70);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 20);
            this.label6.TabIndex = 34;
            this.label6.Text = "Usuario";
            // 
            // txtUserSAP
            // 
            this.txtUserSAP.Location = new System.Drawing.Point(130, 70);
            this.txtUserSAP.Name = "txtUserSAP";
            this.txtUserSAP.Size = new System.Drawing.Size(138, 20);
            this.txtUserSAP.TabIndex = 33;
            // 
            // label7
            // 
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(25, 42);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(100, 20);
            this.label7.TabIndex = 32;
            this.label7.Text = "Base Datos";
            // 
            // txtBDSAP
            // 
            this.txtBDSAP.Location = new System.Drawing.Point(130, 42);
            this.txtBDSAP.Name = "txtBDSAP";
            this.txtBDSAP.Size = new System.Drawing.Size(138, 20);
            this.txtBDSAP.TabIndex = 31;
            // 
            // label8
            // 
            this.label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(24, 16);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(100, 20);
            this.label8.TabIndex = 30;
            this.label8.Text = "Server";
            // 
            // TxtServerSAP
            // 
            this.TxtServerSAP.Location = new System.Drawing.Point(130, 16);
            this.TxtServerSAP.Name = "TxtServerSAP";
            this.TxtServerSAP.Size = new System.Drawing.Size(138, 20);
            this.TxtServerSAP.TabIndex = 29;
            // 
            // ciCompanyList12
            // 
            this.ciCompanyList12.Location = new System.Drawing.Point(20, 295);
            this.ciCompanyList12.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ciCompanyList12.Name = "ciCompanyList12";
            this.ciCompanyList12.Size = new System.Drawing.Size(358, 36);
            this.ciCompanyList12.TabIndex = 4;
            this.ciCompanyList12.Visible = false;
            // 
            // Base
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 331);
            this.Controls.Add(this.ciCompanyList12);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Base";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Base";
            this.Load += new System.EventHandler(this.Base_Load);
            this.Shown += new System.EventHandler(this.Base_Shown);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPass;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtBD;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtServer;
        protected controles.CICompanyList1 ciCompanyList11;
        public System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button buttonversap;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtpwdSAP;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtUserSAP;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtBDSAP;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox TxtServerSAP;
        private controles.CICompanyList1 ciCompanyList12;
    }
}