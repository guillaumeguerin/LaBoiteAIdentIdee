using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AssociationCSV
{
    public partial class Form2 : Form
    {
        public static string text {get; set;}

        public Form2()
        {
            InitializeComponent();
        }

        public Form2(string text2)
        {
            text = text2;
            InitializeComponent();
        }

        public Form2(string text2, string name, string firstname)
        {
            text = text2;
            InitializeComponent();
            this.Text = "Preview - " + firstname + " " + name;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            webBrowser1.DocumentText = text;
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
                text = webBrowser1.DocumentText;
                DialogResult = DialogResult.OK;
        }
    }
}
