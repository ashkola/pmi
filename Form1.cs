using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PMI
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			Config.Get();
			this.textBox1.Text = Config._server;
			this.textBox2.Text = Config._login;
			this.textBox3.Text = Config._password;
			this.textBox4.Text = Config._domain;
			this.textBox5.Text = Config._project;
			this.textBox6.Text = Config._root;
			this.checkBox1.Checked = Config._attachment;
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			
		}

		private void label1_Click(object sender, EventArgs e)
		{
			
		}

		private void button1_Click(object sender, EventArgs e)
		{
			startButton.Enabled = false;
			Config._server = this.textBox1.Text;
			Config._login = this.textBox2.Text;
			Config._password = this.textBox3.Text;
			Config._domain = this.textBox4.Text;
			Config._project = this.textBox5.Text;
			Config._root = this.textBox6.Text;
			Config._attachment = this.checkBox1.Checked;
			Config.Save();
			var main = new Main();
			main.DoWork();
//			main.DoOpenXml();
			startButton.Enabled = true;
		}

	}
}
