namespace PMI
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
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.textBox2 = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.startButton = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.textBox3 = new System.Windows.Forms.TextBox();
			this.textBox4 = new System.Windows.Forms.TextBox();
			this.textBox5 = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.textBox6 = new System.Windows.Forms.TextBox();
			this.checkBox1 = new System.Windows.Forms.CheckBox();
			this.label7 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(96, 6);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(300, 20);
			this.textBox1.TabIndex = 0;
			this.textBox1.Text = "http://stend-qa:8080/qcbin";
			// 
			// textBox2
			// 
			this.textBox2.Location = new System.Drawing.Point(96, 28);
			this.textBox2.Name = "textBox2";
			this.textBox2.Size = new System.Drawing.Size(155, 20);
			this.textBox2.TabIndex = 1;
			this.textBox2.Text = "ashkola";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 13);
			this.label1.TabIndex = 2;
			this.label1.Text = "Server:";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// startButton
			// 
			this.startButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.startButton.Location = new System.Drawing.Point(316, 113);
			this.startButton.Name = "startButton";
			this.startButton.Size = new System.Drawing.Size(80, 23);
			this.startButton.TabIndex = 7;
			this.startButton.Text = "Start";
			this.startButton.UseVisualStyleBackColor = true;
			this.startButton.Click += new System.EventHandler(this.button1_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(12, 31);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Login:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 53);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 13);
			this.label3.TabIndex = 5;
			this.label3.Text = "Password:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(12, 75);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(46, 13);
			this.label4.TabIndex = 6;
			this.label4.Text = "Domain:";
			// 
			// textBox3
			// 
			this.textBox3.Location = new System.Drawing.Point(96, 50);
			this.textBox3.Name = "textBox3";
			this.textBox3.Size = new System.Drawing.Size(155, 20);
			this.textBox3.TabIndex = 2;
			this.textBox3.UseSystemPasswordChar = true;
			// 
			// textBox4
			// 
			this.textBox4.Location = new System.Drawing.Point(96, 72);
			this.textBox4.Name = "textBox4";
			this.textBox4.Size = new System.Drawing.Size(155, 20);
			this.textBox4.TabIndex = 3;
			this.textBox4.Text = "EXTERNAL";
			// 
			// textBox5
			// 
			this.textBox5.Location = new System.Drawing.Point(96, 94);
			this.textBox5.Name = "textBox5";
			this.textBox5.Size = new System.Drawing.Size(155, 20);
			this.textBox5.TabIndex = 4;
			this.textBox5.Text = "Банк222";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(12, 97);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(43, 13);
			this.label5.TabIndex = 10;
			this.label5.Text = "Project:";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(12, 119);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(71, 13);
			this.label6.TabIndex = 11;
			this.label6.Text = "Subject Filter:";
			// 
			// textBox6
			// 
			this.textBox6.Location = new System.Drawing.Point(96, 116);
			this.textBox6.Name = "textBox6";
			this.textBox6.Size = new System.Drawing.Size(155, 20);
			this.textBox6.TabIndex = 5;
			this.textBox6.Text = "^Subject^";
			// 
			// checkBox1
			// 
			this.checkBox1.AutoSize = true;
			this.checkBox1.Location = new System.Drawing.Point(316, 30);
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Size = new System.Drawing.Size(80, 17);
			this.checkBox1.TabIndex = 6;
			this.checkBox1.Text = "Attachment";
			this.checkBox1.UseVisualStyleBackColor = true;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(353, 140);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(43, 13);
			this.label7.TabIndex = 12;
			this.label7.Text = "v.03.14";
			// 
			// Form1
			// 
			this.AcceptButton = this.startButton;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(419, 162);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.checkBox1);
			this.Controls.Add(this.textBox6);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.textBox5);
			this.Controls.Add(this.textBox4);
			this.Controls.Add(this.textBox3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.startButton);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.textBox2);
			this.Controls.Add(this.textBox1);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MaximumSize = new System.Drawing.Size(435, 200);
			this.MinimumSize = new System.Drawing.Size(435, 200);
			this.Name = "Form1";
			this.Text = "HP QC Exporter (TestCases)";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.TextBox textBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button startButton;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBox3;
		private System.Windows.Forms.TextBox textBox4;
		private System.Windows.Forms.TextBox textBox5;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox textBox6;
		private System.Windows.Forms.CheckBox checkBox1;
		private System.Windows.Forms.Label label7;
	}
}

