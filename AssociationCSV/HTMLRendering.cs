using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace AssociationCSV
{
    class HTMLRendering
    {
        public static string buildHtml(Row row)
        {
            string path = Directory.GetCurrentDirectory() + "/Association.html";
            String text = System.IO.File.ReadAllText(path, Encoding.GetEncoding("windows-1252"));

            //TODO sort cells by longest strings.

            for (int i = 0; i < row.getRows()[0].getCells().Count; i++)
            {
                string textToBeReplaced = "##" + row.getRows()[0].getCells()[i].getCell();
                string textForReplacement = row.getRows()[1].getCells()[i].getCell();
                text = text.Replace(textToBeReplaced, textForReplacement);
                //TODO use the longest words first, ex ##ident01Truc then ##id
            }

            text = Regex.Replace(text, @"##[^<\ ]*", "");
            return text;
        }

        public static string buildHtml(List<List<string>> cells)
        {
            string path = Directory.GetCurrentDirectory() + "/Association.html";
            String text = System.IO.File.ReadAllText(path, Encoding.GetEncoding("windows-1252"));

            //TODO sort cells by longest strings.
            cells = reOrderCellsByLongestID(cells);

            for (int i = 0; i < cells[0].Count; i++)
            {
                if (cells[0][i].ToLower() != "art")
                {
                    string textToBeReplaced = "##" + cells[0][i];
                    string textForReplacement = cells[1][i];
                    text = text.Replace(textToBeReplaced, textForReplacement);
                }
                else {
                    Regex.Replace(text, @"##art<", "<");
                }

                //TODO use the longest words first, ex ##ident01Truc then ##id
            }

            text = Regex.Replace(text, @"##[^<\ ]*", "");
            text = text.Replace("<td></td>", "");
            text = text.Replace(" :", ":");
            text = text.Replace("<td></td>", "");
            text = text.Replace("projetsolo1nom", "");
            text = text.Replace("projetsolo2nom", "");
            text = text.Replace("projetsolo3nom", "");
            text = Regex.Replace(text, @"(\ \n)*<td>De:<\/td>(\ |\n)*<td>A:<\/td>(\ |\n)*<td>Traces:<\/td>(\ \n)*", "");

            //text = Regex.Replace(text, "<td>Scènes:</td>(.|\n)*</tr>", "</tr>");
            // Regex are slow as f***

            /*text = Regex.Replace(text, "<table style=\"text-align: left; width: 100px;\" border=\"1\"(.|\n)*cellpadding=\"2\" cellspacing=\"2\">(.|\n)*<tbody>(.|\n)*<tr>(.|\n)*<td colspan=\"4\"><span style=\"font-weight: bold;\"><(.|\n)*</table>", "");
            text = Regex.Replace(text, "<table style=\"text-align: left; width: 100px;\" border=\"1\" cellpadding=\"2\" cellspacing=\"2\">(.|\n)*<tbody>(.|\n)*<tr>(.|\n)*<td></td>(.|\n)*<td></td>(.|\n)*<td>De:</td>(.|\n)*<td></td>", "");*/
            return text;
        }

        public static List<List<string>> reOrderCellsByLongestID(List<List<string>> cells) {
            List<List<string>> result = new List<List<string>>();

            List<string> line1 = new List<string>();
            List<string> line2 = new List<string>();

            while(cells[0].Count > 0) {
                int maxID = getLongestIDInSheet(cells);

                line1.Add(cells[0][maxID]);
                line2.Add(cells[1][maxID]);

                cells[0].RemoveAt(maxID);
                cells[1].RemoveAt(maxID);
            }

            result.Add(line1);
            result.Add(line2);

            return result;
        }

        public static int getLongestIDInSheet(List<List<string>> cells)
        { 
            int longestLength = 0;
            int position = 0;
            for (int i = 0; i < cells[0].Count; i++)
            {
                if (cells[0][i].Length > longestLength)
                {
                    longestLength = cells[0][i].Length;
                    position = i;
                }
            }
            return position;
        }

        public static void generatePDFs(List<List<string>> cells)
        {
            PdfDocument pdf = PdfGenerator.GeneratePdf(getHeader() + getContent(cells) + getFooter(), PageSize.A4);
            string fileName = "CV_" + getFirstName(cells).ToUpper() + getLastName(cells).ToUpper() + ".pdf";
            pdf.Save(fileName);

            /*string fileName = "CV_" + getFirstName(cells).ToUpper() + getLastName(cells).ToUpper() + ".pdf";
            string myHTML = getHeader() + getContent(cells) + getFooter();
            Bitmap bitmap = new Bitmap(790, 1800);
            Graphics g = Graphics.FromImage(bitmap);
            XGraphics xg = XGraphics.FromGraphics(g, new XSize(bitmap.Width, bitmap.Height));
            TheArtOfDev.HtmlRenderer.PdfSharp.HtmlContainer c = new TheArtOfDev.HtmlRenderer.PdfSharp.HtmlContainer();
            c.SetHtml(myHTML);

            PdfDocument pdf = new PdfDocument();
            PdfPage page = new PdfPage();
            XImage img = XImage.FromGdiPlusImage(bitmap);
            pdf.Pages.Add(page);
            XGraphics xgr = XGraphics.FromPdfPage(pdf.Pages[0]);
            c.PerformLayout(xgr);
            c.PerformPaint(xgr);
            xgr.DrawImage(img, 0, 0);
            pdf.Save(fileName);*/
        }

        public static string getFirstName(List<List<string>> cells) {
            for (int i = 0; i < cells[0].Count; i++ )
            {
                if (cells[0][i] == "prenom") {
                    return cells[1][i];
                }
            }
            return "";
        }

        public static string getFirstName(Row sheet)
        {
            for (int i = 0; i < sheet.getRows()[0].getCells().Count; i++)
            {
                if (sheet.getRows()[0].getCells()[i].getCell() == "prenom")
                {
                    return sheet.getRows()[1].getCells()[i].getCell();
                }
            }
            return "";
        }

        public static string getLastName(Row sheet)
        {
            for (int i = 0; i < sheet.getRows()[0].getCells().Count; i++)
            {
                if (sheet.getRows()[0].getCells()[i].getCell() == "nom")
                {
                    return sheet.getRows()[1].getCells()[i].getCell();
                }
            }
            return "Doe";
        }

        public static string getLastName(List<List<string>> cells)
        {
            for (int i = 0; i < cells[0].Count; i++)
            {
                if (cells[0][i] == "nom")
                {
                    return cells[1][i];
                }
            }
            return "Doe";
        }

        public static string getHeader() {
            string header = "<html><head><meta charset='utf-8'><link href='css/bootstrap.min.css' rel='stylesheet'></head><body>";
            return header;
        }

        public static string getContent(List<List<string>> cells)
        {
            string result = "<div class='container'><div class='page-header'><h1>#Prenom #Nom</h1></div>#Genre</br>";
            result = result.Replace("#Nom", getLastName(cells));
            result = result.Replace("#Prenom", getFirstName(cells));
            result = result.Replace("#Genre", getGenre(cells));
            result += getContactInformations(cells);

            result += getRemarque();
            result += getId(cells);
            result += getHobby(cells);
            result += getMusic(cells);
            result += getInformal(cells);
            result += getBand(cells);
            result += getAlone(cells);
            result += getAloneProjects(cells);
            result += getSchool(cells);
            result += getMusicalIdentity(cells);
            result += getSDCMusical(cells);
            result += getFamilly(cells);
            result += getMusicalContext(cells);
            result += getArt(cells);
            result += getArtContext(cells);
            return result;
        }

        public static string getFooter()
        {
            string footer = "</body></html>";
            return footer;
        }

        public static string getGenre(List<List<string>> cells)
        {
            for (int i = 0; i < cells[0].Count; i++)
            {
                if (cells[0][i] == "genre")
                {
                    return cells[1][i];
                }
            }
            return "";
        }

        public static string searchForValue(List<List<string>> cells, string value)
        {
            for (int i = 0; i < cells[0].Count; i++)
            {
                if (cells[0][i] == value)
                {
                    return cells[1][i];
                }
            }
            return "";
        }

        public static string sep(string s)
        {
            int l = s.IndexOf(".");
            if (l > 0)
            {
                return s.Substring(0, l);
            }
            return "";

        }

        public static string getContactInformations(List<List<string>> cells) {
            if (searchForValue(cells, "telephone") != "" || searchForValue(cells, "email") != "" || searchForValue(cells, "postal") != "" || searchForValue(cells, "naissance") != "")
            {
                string res = "<div><h3>Coordonnées</h3><ul >";
                if(searchForValue(cells, "telephone") != "") {
                    res += "<p>Téléphone : "+ searchForValue(cells, "telephone") + "</br></p>";
                }
                if (searchForValue(cells, "email") != "")
                {
                    res += "<p>Email : " + searchForValue(cells, "email") + "</p>";
                }
                if (searchForValue(cells, "postal") != "")
                {
                    try {
                        string str = searchForValue(cells, "postal");
                            res += "<p>Code Postal : " + sep(str) + "</p>";
                    }
                    catch(Exception e) {
                    
                    }
                    
                }
                if (searchForValue(cells, "naissance") != "")
                {
                    res += "<p>Date de naissance : " + searchForValue(cells, "naissance") + "</p>";
                }
                res += "</ul></div></br>";
                return res;
            }
            else
                return "";
        }

        public static string getRemarque() {
            string str = "<h3>Remarque : </h3><p></p>";
            return str;
        }

        public static string getId(List<List<string>> cells)
        {
            if (searchForValue(cells, "id") != "")
            {
                string res = "<h3>Id : " + searchForValue(cells, "id") + "</h3><ul>";
                if (searchForValue(cells, "submitdate") != "")
                {
                    res += "<p>submitdate :" + searchForValue(cells, "submitdate") + "</p>";
                }
                if (searchForValue(cells, "lastpage") != "")
                {
                    res += "<p>lastpage : " + searchForValue(cells, "lastpage") + "</p>";
                }
                if (searchForValue(cells, "startlanguage") != "")
                {
                    res += "<p>startlanguage : " + searchForValue(cells, "startlanguage") + "</p>";
                }
                return res +"</ul>";
            }
            else
                return "";
        }

        public static string getHobby(List<List<string>> cells)
        {
            string res = "";
            if (searchForValue(cells, "art") != "")
            {
                res = "<h3>Art : " + searchForValue(cells, "art") + "</h3>";
            }
            if (searchForValue(cells, "art") != "")
            {
                res = "<h3>Musique : " + searchForValue(cells, "musique") + "</h3>";
            }
            return res;
        }

        public static string getMusic(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "musique") != "")
            {
                result += "<h2>Musique</h2><ul>";
                for (int i = 1; i < 16; i++) {
                    string str1 = "";
                    if (i < 10)
                    {
                        str1 = "0" + i.ToString();
                    }
                    else
                        str1 = i.ToString();
                    if (searchForValue(cells, "instrument[INSTR"+ str1 +"]") != "") {
                        string toolMusic = searchForValue(cells, "instrument[INSTR" + str1 + "]");
                        result += "<h4>" + toolMusic + "</h4><ul>";
                        if (searchForValue(cells, "ctxtinstr"+ str1 +"[groupe]") != "")
                        {
                            result += "<p>ctxtinstr<b>" + toolMusic.Replace(" ", "") + "</b>[groupe] :" + searchForValue(cells, "ctxtinstr" + str1 + "[groupe]") + "</p>";
                        }
                        if (searchForValue(cells, "ctxtinstr" + str1 + "[seul]") != "")
                        {
                            result += "<p>ctxtinstr<b>" + toolMusic.Replace(" ", "") + "</b>[seul] :" + searchForValue(cells, "ctxtinstr" + str1 + "[seul]") + "</p>";
                        }
                        if (searchForValue(cells, "ctxtinstr" + str1 + "[ecole]") != "")
                        {
                            result += "<p>ctxtinstr<b>" + toolMusic.Replace(" ", "") + "</b>[ecole] :" + searchForValue(cells, "ctxtinstr" + str1 + "[ecole]") + "</p>";
                        }
                        if (searchForValue(cells, "ctxtinstr" + str1 + "[informel]") != "")
                        {
                            result += "<p>ctxtinstr<b>" + toolMusic.Replace(" ","") + "</b>[informel] :" + searchForValue(cells, "ctxtinstr" + str1 + "[informel]") + "</p>";
                        }
                        else {
                            result += "Début : " + searchForValue(cells, "tmpsinstr" + str1 + "[SQ001_SQ001]") + " Arrêt : " + searchForValue(cells, "tmpsinstr" + str1 + "[SQ001_SQ002]");
                        }

                        result += "</ul>";
                    }
                }
            }
                return "</ul>"+result;
        }

        public static string getInformal(List<List<string>> cells) {
            string result = "";
            return result;
        }

        public static string getSchool(List<List<string>> cells)
        {
            string result = "";
            for (int i = 1; i < 6; i++ )
            {
                if (searchForValue(cells, "ecole" + i + "nom") != "") {
                    result += "<h4>Ecole : " + searchForValue(cells, "ecole"+i+"nom") + "</h4><ul>";
                    if (searchForValue(cells, "ecole"+i+"type") != "")
                    {
                        result += "<p>ecole" + i + "type : " + searchForValue(cells, "ecole" + i + "type") + "</p>";
                    }
                    result += "<p>De : " + searchForValue(cells, "ecole" + i + "tmps[SQ001_SQ001]") + " A : " + searchForValue(cells, "ecole" + i + "tmps[SQ001_SQ002]</p>");
                    if(searchForValue(cells, "ecole" + i + "traces") != "") {
                        result += "<p>Trace : " + searchForValue(cells, "ecole" + i + "traces") + "</p>";
                    }
                    if(searchForValue(cells, "ecole" + i + "motiv[Apprentissage]") != "") {
                        result += "<p>motiv[Apprentissage] : " + searchForValue(cells, "ecole" + i + "motiv[Apprentissage]") + "</p>";
                    }
                    if(searchForValue(cells, "ecole" + i + "motiv[Composition]") != "") {
                        result += "<p>motiv[Composition] : " + searchForValue(cells, "ecole" + i + "motiv[Composition]") + "</p>";
                    }
                    if(searchForValue(cells, "ecole" + i + "motiv[Improvisation]") != "") {
                        result += "<p>motiv[Improvisation] : " + searchForValue(cells, "ecole" + i + "motiv[Improvisation]") + "</p>";
                    }
                    if(searchForValue(cells, "ecole" + i + "motiv[Reprises]") != "") {
                        result += "<p>motiv[Reprises] : " + searchForValue(cells, "ecole" + i + "motiv[Reprises]") + "</p>";
                    }
                    if (searchForValue(cells, "ecole" + i + "motiv[Plaisir]") != "")
                    {
                        result += "<p>motiv[Plaisir] : " + searchForValue(cells, "ecole" + i + "motiv[Plaisir]") + "</p>";
                    }
                    result += "<p>Instruments : ";
                    for (int k = 1; k < 16; k++) {
                        string str1 = "";
                        if (k < 10)
                        {
                            str1 = "0" + k.ToString();
                        }
                        if(searchForValue(cells, "ecole" + i + "instr[INSTR"+ str1 +"]") == "Oui") {
                            result += " " + searchForValue(cells, "instrument[INSTR" + str1 + "]");
                        }
                    }
                    result += "</ul>";
                }
                
            }
            if (result != "") {
                result = "<h3>Ecoles</h3>" + result;
            }
            return result;
        }

        private static string getFamilly(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "famille") != "")
            {
                result += "<h4> Famille : " + searchForValue(cells, "famille") + "</h4><ul>";
                if (searchForValue(cells, "membrefamille") != "")
                {
                    result += "<p>Membres : " + searchForValue(cells, "membrefamille")+"</p>";
                }
                if (searchForValue(cells, " jeufamille") != "")
                {
                    result += "<p>Jeu Famille : " + searchForValue(cells, " jeufamille")+"</p>";
                }
                result += "</ul>";
            }
            return result;
        }

        private static string getSDCMusical(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "competenceMusique[1]") != "" || searchForValue(cells, "competenceMusique[2]") != "" || searchForValue(cells, "competenceMusique[3]") != "")
            {
                result += "<h3>SDC musical</h3><ul>";
            }
            if (searchForValue(cells, "competenceMusique[1]") != "")
            {
                result += "<p>competenceMusique[1] : " + searchForValue(cells, "competenceMusique[1]") + "</p>";
            }
            if (searchForValue(cells, "competenceMusique[2]") != "")
            {
                result += "<p>competenceMusique[2] : " + searchForValue(cells, "competenceMusique[2]") + "</p>";
            }
            if (searchForValue(cells, "competenceMusique[3]") != "")
            {
                result += "<p>competenceMusique[3] : " + searchForValue(cells, "competenceMusique[3]") + "</p>";
            }
            if(result != ""){
                result += "</ul>";
            }
            return result;
        }

        private static string getMusicalIdentity(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "identitMusique[1]") != "" || searchForValue(cells, "identitMusique[2]") != "" || searchForValue(cells, "identitMusique[3]") != "")
            {
                result += "<h3>Identite musicale</h3><ul>";
            }
            if (searchForValue(cells, "identitMusique[1]") != "")
            {
                result += "<p>identitMusique[1] : " + searchForValue(cells, "identitMusique[1]") + "</p>";
            }
            if (searchForValue(cells, "identitMusique[2]") != "")
            {
                result += "<p>identitMusique[2] : " + searchForValue(cells, "identitMusique[2]") + "</p>";
            }
            if (searchForValue(cells, "identitMusique[3]") != "")
            {
                result += "<p>identitMusique[3] : " + searchForValue(cells, "identitMusique[3]") + "</p>";
            }
            if(result != "") {
                result += "</ul>";
            }
            return result;
        }

        private static string getAloneProjects(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "projetsolo1nom") != "")
            {
                result += "<h3>Projets</h3><ul>";
            }
            for (int i = 1; i < 4; i++ )
            {
                if (searchForValue(cells, "projetsolo"+i+"nom") != "")
                {
                    result += "<h4>" + searchForValue(cells, "projetsolo" + i + "nom") + "</h4>";
                    result += "<p>De " + searchForValue(cells, "projetsolo0" + i + "tmps[SQ001_SQ001]") + " A " + searchForValue(cells, "projetsolo0" + i + "tmps[SQ001_SQ002]")  + "</p>";
                }

                if(searchForValue(cells, "projetsolo"+i+"motiv[Apprentissage]") != "" || searchForValue(cells, "projetsolo" + i + "motiv[Composition]") != "" || searchForValue(cells, "projetsolo" + i + "motiv[Improvisation]") != "" || searchForValue(cells, "projetsolo" + i + "motiv[Reprises]") != "" || searchForValue(cells, "projetsolo" + i + "motiv[Plaisir]") != "") {
                    result += "<h4>Motivations</h4><ul>";
                    if (searchForValue(cells, "projetsolo" + i + "motiv[Apprentissage]") != "")
                    {
                        result += "<p>projetsolo" + i + "motiv[Apprentissage]" + searchForValue(cells, "projetsolo" + i + "motiv[Apprentissage]") + "</p>";
                    }
                    if (searchForValue(cells, "projetsolo" + i + "motiv[Composition]") != "")
                    {
                        result += "<p>projetsolo" + i + "motiv[Composition]" + searchForValue(cells, "projetsolo" + i + "motiv[Composition]") + "</p>";
                    }
                    if (searchForValue(cells, "projetsolo" + i + "motiv[Improvisation]") != "")
                    {
                        result += "<p>projetsolo" + i + "motiv[Improvisation]" + searchForValue(cells, "projetsolo" + i + "motiv[Improvisation]") + "</p>";
                    }
                    if (searchForValue(cells, "projetsolo" + i + "motiv[Reprises]") != "")
                    {
                        result += "<p>projetsolo" + i + "motiv[Reprises]" + searchForValue(cells, "projetsolo" + i + "motiv[Reprises]") + "</p>";
                    }
                    if (searchForValue(cells, "projetsolo" + i + "motiv[Plaisir]") != "")
                    {
                        result += "<p>projetsolo" + i + "motiv[Plaisir]" + searchForValue(cells, "projetsolo" + i + "motiv[Plaisir]") + "</p>";
                    }
                    result += "</ul>";
                }

                for (int l = 1; l < 21; l++)
                {
                    string str1 = "";
                    if (l < 10)
                    {
                        str1 = "0" + l.ToString();
                    }
                    if (searchForValue(cells, "projetsolo" + i + "[INSTR" + str1 + "]") == "Oui")
                    {
                        result += " " + searchForValue(cells, "instrument[INSTR" + str1 + "]");
                    }
                }




                if (searchForValue(cells, "projetsolo" + i + "scene") != "")
                {
                    result += "<p>Scene " + searchForValue(cells, "projetsolo" + i + "scene") + "</p>";
                }
                if (searchForValue(cells, "projetsolo" + i + "traces") != "")
                {
                    result += "<p>Traces" + searchForValue(cells, "projetsolo" + i + "traces") + "</p>";
                }
                result += "</ul>";
            }
            return result;
        }

        private static string getAlone(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "seulgenre") != "")
            {
                result += "<h3>Seul</h3><ul> <p> Genre :" + searchForValue(cells, "seulgenre") + "</p>";
            }
            if (searchForValue(cells, "seultraces") != "")
            {
                result += "<p>Traces :" + searchForValue(cells, "seultraces") + "</p>";
            }
            if (searchForValue(cells, "seulmotiv[Apprentissage]") != "" || searchForValue(cells, "seulmotiv[Composition]") != "" || searchForValue(cells, "seulmotiv[Improvisation]") != "" || searchForValue(cells, "seulmotiv[Reprises]") != "" || searchForValue(cells, "seulmotiv[Plaisir]") != "")
            {
                result += "<h4>Motivations</h4><ul>";
                if (searchForValue(cells, "seulmotiv[Apprentissage]") != "")
                {
                    result += "<p>seulmotiv[Apprentissage] : " + searchForValue(cells, "seulmotiv[Apprentissage]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv[Composition]") != "")
                {
                    result += "<p>seulmotiv[Composition] : " + searchForValue(cells, "seulmotiv[Composition]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv[Improvisation]") != "")
                {
                    result += "<p>seulmotiv[Improvisation] : " + searchForValue(cells, "seulmotiv[Improvisation]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv[Reprises]") != "")
                {
                    result += "<p>seulmotiv[Reprises] : " + searchForValue(cells, "seulmotiv[Reprises]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv[Plaisir]") != "")
                {
                    result += "<p>seulmotiv[Plaisir] : " + searchForValue(cells, "seulmotiv[Plaisir]") + "</p>";
                }

                //??
                if (searchForValue(cells, "seulmotiv2[Motiv01]") != "")
                {
                    result += "<p>seulmotiv2[Motiv01] : " + searchForValue(cells, "seulmotiv2[Motiv01]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv2[Motiv02]") != "")
                {
                    result += "<p>seulmotiv2[Motiv02] : " + searchForValue(cells, "seulmotiv2[Motiv02]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv2[Motiv03]") != "")
                {
                    result += "<p>seulmotiv2[Motiv03] : " + searchForValue(cells, "seulmotiv2[Motiv03]") + "</p>";
                }
                if (searchForValue(cells, "seulmotiv2[Motiv04]") != "")
                {
                    result += "<p>seulmotiv2[Motiv04] : " + searchForValue(cells, "seulmotiv2[Motiv04]") + "</p>";
                }
                result += "<p>";
                for (int i = 1; i < 21; i++ )
                {
                    string str1 = "";
                    if (i < 10)
                    {
                        str1 = "0" + i.ToString();
                    }
                    if (searchForValue(cells, "seulinstr[INSTR" + str1 + "]") == "Oui")
                    {
                        result += " " + searchForValue(cells, "instrument[INSTR" + str1 + "]");
                    }
                }
                if (searchForValue(cells, "seulinstr[other]") != "")
                {
                    result += "seulinstr[other] : " + searchForValue(cells, "seulinstr[other]");
                }
                if (searchForValue(cells, "seulMAO") != "")
                {
                    result += "seulMAO" + searchForValue(cells, "seulMAO");
                }
                result += "</p>";
            }
            return "</ul>"+result;
        }

        private static string getBand(List<List<string>> cells)
        {
            string result = "";
            for (int i = 1; i < 11; i++ )
            {
                string str1 = "";
                if (i < 10)
                {
                    str1 = "0" + i.ToString();
                }
                if (searchForValue(cells, "groupe"+str1+"nom") != "")
                {
                    result += "<h4>Groupe : "+ searchForValue(cells, "groupe"+str1+"nom") +"</h4>";
                    result += "<p>De : " + searchForValue(cells, "groupe" + str1 + "tmps[SQ001_SQ001]") + " A : " + searchForValue(cells, "groupe" + str1 + "tmps[SQ001_SQ002]") + " </p>";
                    if (searchForValue(cells, "groupe" + str1 + "nb") != "")
                    {
                        result += "<p>Groupe nb : " + sep(searchForValue(cells, "groupe" + str1 + "nb")) + "</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "genre") != "")
                    {
                        result += "<p>Genre : " + searchForValue(cells, "groupe" + str1 + "genre") + "</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv[Apprentissage]") != "" || searchForValue(cells, "groupe" + str1 + "motiv[CompositionSQ002]") != "" || searchForValue(cells, "groupe" + str1 + "motiv[Improvisation]") != "" || searchForValue(cells, "groupe" + str1 + "motiv[Reprises]") != "" || searchForValue(cells, "groupe" + str1 + "motiv[Plaisir]") != "" || searchForValue(cells, "groupe" + str1 + "motiv2") != "")
                    {
                        result += "<h5>Motivations : </h5>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv[Apprentissage]") != "")
                    {
                        result += "<p>groupe" + str1 + "motiv[Apprentissage] : " + searchForValue(cells, "groupe" + str1 + "motiv[Apprentissage]") +"</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv[CompositionSQ002]") != "")
                    {
                        result += "<p>groupe" + str1 + "motiv[CompositionSQ002] : " + searchForValue(cells, "groupe" + str1 + "motiv[CompositionSQ002]") + "</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv[Improvisation]") != "")
                    {
                        result += "<p>groupe" + str1 + "motiv[Improvisation] : " + searchForValue(cells, "groupe" + str1 + "motiv[Improvisation]") + "</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv[Reprises]") != "")
                    {
                        result += "<p>groupe" + str1 + "motiv[Reprises] : " + searchForValue(cells, "groupe" + str1 + "motiv[Reprises]") + "</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv[Plaisir]") != "")
                    {
                        result += "<p>groupe" + str1 + "motiv[Plaisir] : " + searchForValue(cells, "groupe" + str1 + "motiv[Plaisir]") + "</p>";
                    }
                    if (searchForValue(cells, "groupe" + str1 + "motiv2") != "")
                    {
                        result += "<p>groupe" + str1 + "motiv2 : " + searchForValue(cells, "groupe" + str1 + "motiv2") + "</p>";
                    }
                   

                    result += "<p>Instruments : </p><p>";
                    /*for (int j = 1; j < 16; j++ )
                    {
                        string str2 = "";
                        if (j < 10)
                        {
                            str2 = "0" + j.ToString();
                        }
                        if (searchForValue(cells, "groupe" + str1 + "[INSTR"+str2+"]") != "")
                        {
                            result += searchForValue(cells, "groupe" + str1 + "[INSTR" + str2 + "]") + " ";
                        }
                        result += searchForValue(cells, "groupe" + str1 + "[other]") + " " + searchForValue(cells, "groupe" + str1 + "MAO");
                    }*/
                    for (int j = 1; j < 16; j++)
                    {
                        string str2 = "";
                        if (j < 10)
                        {
                            str2 = "0" + j.ToString();
                        }
                        if (searchForValue(cells, "groupe" + str1 + "[INSTR" + str2 + "]") != "")
                        {
                            result += searchForValue(cells, "instrument[INSTR" + str2 + "]") + " ";
                        }
                    }
                    if (searchForValue(cells, "groupe" + str1 + "[other]") != "")
                    {
                        result += "groupe" + str1 + "[other] ";
                    }

                    if (searchForValue(cells, "groupe" + str1 + "MAO") != "")
                    {
                        result += "MAO ";
                    }
                    result += "</p>";
                    if (searchForValue(cells, "groupe" + str1 + "grpinstr") != "")
                    {
                        result += "Instruments du groupe : " + searchForValue(cells, "groupe" + str1 + "grpinstr");
                    }
                    if (searchForValue(cells, "groupe" + str1 + "scenes") != "")
                    {
                        result += "Scenes : " + searchForValue(cells, "groupe" + str1 + "scenes");
                    }
                    if (searchForValue(cells, "groupe" + str1 + "traces") != "")
                    {
                        result += "Traces : " + searchForValue(cells, "groupe" + str1 + "traces");
                    }
                }
            }
            return result;
        }

        private static string getArtContext(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "manqueart") != "")
            {
                result += "<h4>Contextes artistiques manquants :</h4><ul><p>manqueart : " + searchForValue(cells, "manqueart") + "</p>";
            }
            result += "<p>" + searchForValue(cells, "manqueartTXT") + "</p>";
            if (searchForValue(cells, "manqueartmotiv[Apprentissage]") != "")
            {
                result += "<p>manqueartmotiv[Apprentissage] : " + searchForValue(cells, "manqueartmotiv[Apprentissage]") + "</p>";
            }
            if (searchForValue(cells, "manqueartmotiv[Creation]") != "")
            {
                result += "<p>manqueartmotiv[Creation] : " + searchForValue(cells, "manqueartmotiv[Creation]") + "</p>";
            }
            if (searchForValue(cells, "manqueartmotiv[Libre]") != "")
            {
                result += "<p>manqueartmotiv[Libre] : " + searchForValue(cells, "manqueartmotiv[Libre]") + "</p>";
            }

            if (searchForValue(cells, "manqueartmotiv[Reprises]") != "")
            {
                result += "<p>manqueartmotiv[Reprises] : " + searchForValue(cells, "manqueartmotiv[Reprises]") + "</p>";
            }
            if (searchForValue(cells, "manqueartmotiv[Plaisir]") != "")
            {
                result += "<p>manqueartmotiv[Plaisir] : " + searchForValue(cells, "manqueartmotiv[Plaisir]") + "</p>";
            }
            if (searchForValue(cells, "manqueartmotiv2") != "")
            {
                result += "<p>manqueartmotiv2 : " + searchForValue(cells, "manqueartmotiv2") + "</p>";
            }
            if(result != "") {
                result += "</ul>";
            }
            return result;
        }

        private static string getArt(List<List<string>> cells)
        {
            string result = "";
            for (int j = 1; j < 11; j++)
            {
                string str2 = "";
                if (j < 10)
                {
                    str2 = "0" + j.ToString();
                }
                if (searchForValue(cells, "art" + str2) != "")
                {
                    result += "<h3>" + searchForValue(cells, "art" + str2) + " </h3>";
                }
                if (searchForValue(cells, "art" + str2 + "tmps[SQ001_SQ001]") != "" || searchForValue(cells, "art" + str2 + "tmps[SQ001_SQ002]") != "")
                {
                    result += "<p>De : " + searchForValue(cells, "art" + str2 + "tmps[SQ001_SQ001]") + " A :" + searchForValue(cells, "art" + str2 + "tmps[SQ001_SQ002]") + "</p>";

                }
                if (searchForValue(cells, "art" + str2 + "parcours") != "")
                {
                    result += "<p>" + searchForValue(cells, "art" + str2 + "parcours") + "</p>";
                }
            }
            return result;
        }

        private static string getMusicalContext(List<List<string>> cells)
        {
            string result = "";
            if (searchForValue(cells, "manquecontexte") != "")
            {
                result += "<h3>Manque contexte musical</h3><ul><p>manquecontexte : " + searchForValue(cells, "manquecontexte") + "</p>";
            }
            if (searchForValue(cells, "manquecontexteTXT") != "")
            {
                result += "<p>"+ searchForValue(cells, "manquecontexteTXT") + "</p>";
            }
            if (searchForValue(cells, "manquecontexteinstr1[INSTR01]") != "" || searchForValue(cells, "manquecontexteinstr1[INSTR02]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR03]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR04]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR05]") != "" 
                ||searchForValue(cells, "manquecontexteinstr1[INSTR06]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR07]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR08]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR09]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR10]") != "" 
                ||searchForValue(cells, "manquecontexteinstr1[INSTR11]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR12]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR13]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR14]") != "" ||searchForValue(cells, "manquecontexteinstr1[INSTR15]") != "")
            {
                result += "<h4>Instruments</h4><p>";
            }
            
            for (int j = 1; j < 16; j++)
            {
                string str2 = "";
                if (j < 10)
                {
                    str2 = "0" + j.ToString();
                }
                if (searchForValue(cells, "manquecontexteinstr1[INSTR"+str2+"]") != "")
                {
                    result += "manquecontexteinstr1[INSTR" + str2 + "] : " + searchForValue(cells, "manquecontexteinstr1[INSTR" + str2 + "]") + " ";
                }
            }
            if (searchForValue(cells, "manquecontexteinstr1[other]") != "")
            {
                result += "manquecontexteinstr1[other] : " + searchForValue(cells, "manquecontexteinstr1[other]");
            }
            if (searchForValue(cells, "manquecontexteMAO") != "")
            {
                result += " MAO : " + searchForValue(cells, "manquecontexteMAO");
            }
            result += "</p>";
            if (searchForValue(cells, "manquecontextemotiv[Apprentissage]") != "" || searchForValue(cells, "manquecontextemotiv[CompositionSQ002]") != "" || searchForValue(cells, "manquecontextemotiv[Improvisation]") != "" || 
                searchForValue(cells, "manquecontextemotiv[Reprises]") != "" || 
                searchForValue(cells, "manquecontextemotiv[Plaisir]") != "" || 
                searchForValue(cells, "manquecontextemotiv2") != "")
            {
                result += "Motivations : ";
                if (searchForValue(cells, "manquecontextemotiv[Apprentissage]") != "")
                {
                    result += "<p>manquecontextemotiv[Apprentissage] : " + searchForValue(cells, "manquecontextemotiv[Apprentissage]")+ "</p>";
                }
                if (searchForValue(cells, "manquecontextemotiv[CompositionSQ002]") != "")
                {
                    result += "<p>manquecontextemotiv[CompositionSQ002] : " + searchForValue(cells, "manquecontextemotiv[CompositionSQ002]")+ "</p>";
                }
                if (searchForValue(cells, "manquecontextemotiv[Improvisation]") != "")
                {
                    result += "<p>manquecontextemotiv[Improvisation] : " + searchForValue(cells, "manquecontextemotiv[Improvisation]")+ "</p>";
                }
                if (searchForValue(cells, "manquecontextemotiv[Reprises]") != "")
                {
                    result += "<p>manquecontextemotiv[Reprises] : " + searchForValue(cells, "manquecontextemotiv[Reprises]") + "</p>";
                }
                if (searchForValue(cells, "manquecontextemotiv[Plaisir]") != "")
                {
                    result += "<p>manquecontextemotiv[Plaisir] : " + searchForValue(cells, "manquecontextemotiv[Plaisir]") + "</p>";
                }
                if(searchForValue(cells, "manquecontextemotiv2") != "") {
                    result += "<p>manquecontextemotiv2 : " + searchForValue(cells, "manquecontextemotiv2") + "</p>";
                }

            }
            result += "</p></ul>";
            return result;
        }
    }
}
