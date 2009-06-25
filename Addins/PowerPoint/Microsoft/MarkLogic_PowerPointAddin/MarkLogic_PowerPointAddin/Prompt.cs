using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MarkLogic_PowerPointAddin
{
    public partial class Prompt : Form
    {
        public string pfilename = "";
        public Prompt()
        {
            InitializeComponent();
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("SAVING "+textBox1.Text);
            pfilename = textBox1.Text;
            
          //  if( this.ShowDialog() == DialogResult.OK && textBox1.Text.Length > 0)  
            //{
              //  MessageBox.Show("RETURNING "+textBox1.Text);
           // }


            this.Close();

            
        }

        private void Prompt_Load(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("CANCELING");
            this.Close();
        }
    }
}
