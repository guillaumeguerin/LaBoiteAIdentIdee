using PdfSharp;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace AssociationCSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Waiting for Drag and Drop";
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[]) e.Data.GetData(DataFormats.FileDrop, false);
            toolStripStatusLabel1.Text = "File Received";
            int i=1;
            foreach(string file in files) {
                toolStripStatusLabel1.Text = "Processing file " + i + " of " + files.Length;
                CsvProcessing.processCsv(file);
                if (files.Length > 0)
                {
                    toolStripProgressBar1.Value = i / files.Length * 100;
                }
                
                i++;
            }

            toolStripStatusLabel1.Text = "Done.";
        }


    }
}
