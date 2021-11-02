using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class SapphireDialogBox : Form
    {
        public string methodstr;
        public SapphireDialogBox()
        {
            //exists as a dialog box for the input of method name.
            InitializeComponent();

            this.AcceptButton = button1;
        }
        public SapphireDialogBox(string s)
        {
            InitializeComponent();
            textBox1.Text = s;
            this.AcceptButton = button1;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            methodstr = this.textBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK ;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
