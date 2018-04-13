using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CCS_System
{
    public partial class CalcProgress : Form
    {
        public CalcProgress()
        {
            InitializeComponent();
        }

        public TextBox textBox
        {
            get { return this.textBox1; }
            set { this.textBox1 = value; }
        }

        public ProgressBar progressBar
        {
            get { return this.progressBar1; }
            set { this.progressBar1 = value; }
        }

        public Label label
        {
            get { return this.label1; }
            set { this.label1 = value; }
        }

        public void startTimer()
        {
            this.timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(this.progressBar1.Value == 100)
            {
                this.timer1.Stop();
                enabledCloseButton();
                //this.Close();
            }
        }

        public void enabledCloseButton()
        {
            this.button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
