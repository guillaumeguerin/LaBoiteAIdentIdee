using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AssociationCSV
{
    public partial class PreviewForm : Form
    {
        public static string text {get; set;}

        public PreviewForm()
        {
            InitializeComponent();
        }

        public PreviewForm(string text2)
        {
            text = text2;
            InitializeComponent();
        }

        public PreviewForm(string text2, string name, string firstname)
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
                text = Regex.Replace(text, "<textarea id=\"presentation\"(.)*textarea>", webBrowser1.Document.GetElementById("presentation").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"oubliInstr\"(.)*textarea>", webBrowser1.Document.GetElementById("oubliInstr").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"coursParticulierEcole\"(.)*textarea>", webBrowser1.Document.GetElementById("coursParticulierEcole").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"enseignementAnimation\"(.)*textarea>", webBrowser1.Document.GetElementById("enseignementAnimation").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"techniqueAutreMAO\"(.)*textarea>", webBrowser1.Document.GetElementById("techniqueAutreMAO").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"autreContexte\"(.)*textarea>", webBrowser1.Document.GetElementById("autreContexte").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"sonGenere\"(.)*textarea>", webBrowser1.Document.GetElementById("sonGenere").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"redemandeArt\"(.)*textarea>", webBrowser1.Document.GetElementById("redemandeArt").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"atelierParticipation\"(.)*textarea>", webBrowser1.Document.GetElementById("atelierParticipation").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"domaineExplorationNonAboutie\"(.)*textarea>", webBrowser1.Document.GetElementById("domaineExplorationNonAboutie").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"opportunitesActivitesArtistiques\"(.)*textarea>", webBrowser1.Document.GetElementById("opportunitesActivitesArtistiques").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"hobbies\"(.)*textarea>", webBrowser1.Document.GetElementById("hobbies").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"projetProfessionel\"(.)*textarea>", webBrowser1.Document.GetElementById("projetProfessionel").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"paiement\"(.)*textarea>", webBrowser1.Document.GetElementById("paiement").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"materiel\"(.)*textarea>", webBrowser1.Document.GetElementById("materiel").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"ideeAtelier\"(.)*textarea>", webBrowser1.Document.GetElementById("ideeAtelier").GetAttribute("value"));
                text = Regex.Replace(text, "<textarea id=\"proutprout\"(.)*textarea>", webBrowser1.Document.GetElementById("proutprout").GetAttribute("value"));

                text = Regex.Replace(text, "<input id=\"emploisDuTempsDebut\"(.)*\">", webBrowser1.Document.GetElementById("emploisDuTempsDebut").GetAttribute("value"));
                text = Regex.Replace(text, "<input id=\"emploisDuTempsFin\"(.)*\">", webBrowser1.Document.GetElementById("emploisDuTempsFin").GetAttribute("value"));
                DialogResult = DialogResult.OK;
        }
    }
}
