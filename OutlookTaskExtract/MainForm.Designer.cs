namespace OutlookTaskExtract
{
    partial class MainForm
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
      this.Start = new System.Windows.Forms.Button();
      this.AutoDiscoverURL = new System.Windows.Forms.CheckBox();
      this.ExchangeServerURL = new System.Windows.Forms.TextBox();
      this.Username = new System.Windows.Forms.TextBox();
      this.label2 = new System.Windows.Forms.Label();
      this.label3 = new System.Windows.Forms.Label();
      this.Password = new System.Windows.Forms.TextBox();
      this.label4 = new System.Windows.Forms.Label();
      this.Domain = new System.Windows.Forms.TextBox();
      this.label1 = new System.Windows.Forms.Label();
      this.ExchangeVersions = new System.Windows.Forms.ComboBox();
      this.PanelReadme = new System.Windows.Forms.Panel();
      this.groupBox2 = new System.Windows.Forms.GroupBox();
      this.label5 = new System.Windows.Forms.Label();
      this.groupBox1 = new System.Windows.Forms.GroupBox();
      this.statusStrip1 = new System.Windows.Forms.StatusStrip();
      this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
      this.PanelReadme.SuspendLayout();
      this.groupBox2.SuspendLayout();
      this.groupBox1.SuspendLayout();
      this.statusStrip1.SuspendLayout();
      this.SuspendLayout();
      // 
      // Start
      // 
      this.Start.Location = new System.Drawing.Point(11, 271);
      this.Start.Name = "Start";
      this.Start.Size = new System.Drawing.Size(75, 23);
      this.Start.TabIndex = 0;
      this.Start.Text = "Start";
      this.Start.UseVisualStyleBackColor = true;
      this.Start.Click += new System.EventHandler(this.extract_Click);
      // 
      // AutoDiscoverURL
      // 
      this.AutoDiscoverURL.AutoSize = true;
      this.AutoDiscoverURL.Checked = true;
      this.AutoDiscoverURL.CheckState = System.Windows.Forms.CheckState.Checked;
      this.AutoDiscoverURL.Location = new System.Drawing.Point(313, 117);
      this.AutoDiscoverURL.Name = "AutoDiscoverURL";
      this.AutoDiscoverURL.Size = new System.Drawing.Size(118, 17);
      this.AutoDiscoverURL.TabIndex = 3;
      this.AutoDiscoverURL.Text = "Auto Discover URL";
      this.AutoDiscoverURL.UseVisualStyleBackColor = true;
      this.AutoDiscoverURL.CheckedChanged += new System.EventHandler(this.AutoDiscoverURL_CheckedChanged);
      // 
      // ExchangeServerURL
      // 
      this.ExchangeServerURL.Location = new System.Drawing.Point(123, 117);
      this.ExchangeServerURL.Name = "ExchangeServerURL";
      this.ExchangeServerURL.Size = new System.Drawing.Size(166, 20);
      this.ExchangeServerURL.TabIndex = 4;
      // 
      // Username
      // 
      this.Username.Location = new System.Drawing.Point(123, 28);
      this.Username.Name = "Username";
      this.Username.Size = new System.Drawing.Size(121, 20);
      this.Username.TabIndex = 5;
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(25, 31);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(58, 13);
      this.label2.TabIndex = 6;
      this.label2.Text = "User name";
      // 
      // label3
      // 
      this.label3.AutoSize = true;
      this.label3.Location = new System.Drawing.Point(25, 61);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(53, 13);
      this.label3.TabIndex = 8;
      this.label3.Text = "Password";
      // 
      // Password
      // 
      this.Password.Location = new System.Drawing.Point(123, 58);
      this.Password.Name = "Password";
      this.Password.PasswordChar = '*';
      this.Password.Size = new System.Drawing.Size(121, 20);
      this.Password.TabIndex = 7;
      // 
      // label4
      // 
      this.label4.AutoSize = true;
      this.label4.Location = new System.Drawing.Point(25, 91);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(43, 13);
      this.label4.TabIndex = 10;
      this.label4.Text = "Domain";
      // 
      // Domain
      // 
      this.Domain.Location = new System.Drawing.Point(123, 88);
      this.Domain.Name = "Domain";
      this.Domain.Size = new System.Drawing.Size(166, 20);
      this.Domain.TabIndex = 9;
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(24, 29);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(42, 13);
      this.label1.TabIndex = 14;
      this.label1.Text = "Version";
      // 
      // ExchangeVersions
      // 
      this.ExchangeVersions.DisplayMember = "Exchange 2010";
      this.ExchangeVersions.FormattingEnabled = true;
      this.ExchangeVersions.Items.AddRange(new object[] {
            "Exchange 2010",
            "Exchange 2010 SP1",
            "Exchange 2010 SP2 "});
      this.ExchangeVersions.Location = new System.Drawing.Point(122, 26);
      this.ExchangeVersions.Name = "ExchangeVersions";
      this.ExchangeVersions.Size = new System.Drawing.Size(121, 21);
      this.ExchangeVersions.TabIndex = 13;
      this.ExchangeVersions.ValueMember = "Exchange 2010";
      // 
      // PanelReadme
      // 
      this.PanelReadme.Controls.Add(this.groupBox2);
      this.PanelReadme.Controls.Add(this.Start);
      this.PanelReadme.Location = new System.Drawing.Point(12, 12);
      this.PanelReadme.Name = "PanelReadme";
      this.PanelReadme.Size = new System.Drawing.Size(452, 419);
      this.PanelReadme.TabIndex = 15;
      // 
      // groupBox2
      // 
      this.groupBox2.Controls.Add(this.label5);
      this.groupBox2.Controls.Add(this.label2);
      this.groupBox2.Controls.Add(this.ExchangeServerURL);
      this.groupBox2.Controls.Add(this.label4);
      this.groupBox2.Controls.Add(this.AutoDiscoverURL);
      this.groupBox2.Controls.Add(this.Username);
      this.groupBox2.Controls.Add(this.Domain);
      this.groupBox2.Controls.Add(this.Password);
      this.groupBox2.Controls.Add(this.label3);
      this.groupBox2.Location = new System.Drawing.Point(3, 92);
      this.groupBox2.Name = "groupBox2";
      this.groupBox2.Size = new System.Drawing.Size(446, 161);
      this.groupBox2.TabIndex = 3;
      this.groupBox2.TabStop = false;
      this.groupBox2.Text = "Authenticate";
      // 
      // label5
      // 
      this.label5.AutoSize = true;
      this.label5.Location = new System.Drawing.Point(25, 121);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(80, 13);
      this.label5.TabIndex = 11;
      this.label5.Text = "Exchange URL";
      // 
      // groupBox1
      // 
      this.groupBox1.Controls.Add(this.label1);
      this.groupBox1.Controls.Add(this.ExchangeVersions);
      this.groupBox1.Location = new System.Drawing.Point(14, 16);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new System.Drawing.Size(446, 77);
      this.groupBox1.TabIndex = 4;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "Exchange Version";
      // 
      // statusStrip1
      // 
      this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
      this.statusStrip1.Location = new System.Drawing.Point(0, 330);
      this.statusStrip1.Name = "statusStrip1";
      this.statusStrip1.Size = new System.Drawing.Size(477, 22);
      this.statusStrip1.TabIndex = 16;
      this.statusStrip1.Text = "statusStrip1";
      // 
      // toolStripStatusLabel1
      // 
      this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
      this.toolStripStatusLabel1.Size = new System.Drawing.Size(103, 17);
      this.toolStripStatusLabel1.Text = "Not authenticated";
      // 
      // MainForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(477, 352);
      this.Controls.Add(this.statusStrip1);
      this.Controls.Add(this.groupBox1);
      this.Controls.Add(this.PanelReadme);
      this.Name = "MainForm";
      this.Text = "Outlook Task Export";
      this.PanelReadme.ResumeLayout(false);
      this.groupBox2.ResumeLayout(false);
      this.groupBox2.PerformLayout();
      this.groupBox1.ResumeLayout(false);
      this.groupBox1.PerformLayout();
      this.statusStrip1.ResumeLayout(false);
      this.statusStrip1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Start;
        public System.Windows.Forms.CheckBox AutoDiscoverURL;
        public System.Windows.Forms.TextBox ExchangeServerURL;
        private System.Windows.Forms.TextBox Username;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Domain;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox ExchangeVersions;
        private System.Windows.Forms.Panel PanelReadme;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
    }
}

