using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AssociationCSV
{
    class HTMLRendering
    {

        public static string buildHtml(List<List<string>> cells)
        {
            string path = Directory.GetCurrentDirectory() + "/Association.html";
            String text = System.IO.File.ReadAllText(path, Encoding.GetEncoding("windows-1252"));

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

        public static string getFirstName(List<List<string>> cells) {
            for (int i = 0; i < cells[0].Count; i++ )
            {
                if (cells[0][i] == "prenom") {
                    return cells[1][i];
                }
            }
            return "";
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

    }
}
