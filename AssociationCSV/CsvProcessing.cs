using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AssociationCSV
{
    class CsvProcessing
    {
        internal static void processCsv(string path)
        {
            if (checkFileExtension(path))
            {
                try
                {
                    string[] fileContent = readLines(path);
                    List<List<string>> cells = getCells(fileContent);
                    List<List<string>>[] sheets = separateUsersInSheets(cells);
                    foreach (List<List<string>> mySheet in sheets)
                    {
                        processSheet(mySheet);
                        string name = HTMLRendering.getLastName(mySheet).ToUpper();
                        string firstname = HTMLRendering.getFirstName(mySheet).ToUpper();
                        string fileName = "CV_" + firstname + name;
                        string text = HTMLRendering.buildHtml(mySheet);
                        PreviewForm previewForm = new PreviewForm(text, name, firstname);
                        DialogResult dr = previewForm.ShowDialog();
                        if(dr == DialogResult.OK) {
                            text = PreviewForm.text;
                        }
                        string path2 = Directory.GetCurrentDirectory() + "\\" + fileName + ".html";
                        File.WriteAllText(Directory.GetCurrentDirectory() + "\\" + fileName + ".html", text, Encoding.GetEncoding("windows-1252"));

                        string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        Directory.SetCurrentDirectory(exeDir);
                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                        startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                        startInfo.FileName = "cmd.exe";
                        startInfo.Arguments = "/C phantomjs.exe pdf.js " + fileName + ".html " + fileName + ".pdf";
                        process.StartInfo = startInfo;
                        process.StartInfo.WorkingDirectory = exeDir; 
                        process.Start();
                        process.WaitForExit();
                    }

                }
                catch (IOException e)
                {
                    MessageBox.Show("Une erreur est survenue lors de la lecture du fichier .csv.\nVeuillez vous assurer que celui-ci existe ou qu'il ne soit pas déjà ouvert.",
                        "Erreur sur le fichier", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else {
                MessageBox.Show("Le format du fichier :" + path + " est incorrect, le format attendu est '*.csv'");
            }
        }

        private static List<List<string>> processSheet(List<List<string>> cells)
        {
            cells = cleanCells(cells);
            cells = emptyCells(cells);
            cells = removeCells(cells);
            return cells;
        } 

        private static bool checkFileExtension(string fileName) {
            Regex regex = new Regex(@".*\.csv");
            Match match = regex.Match(fileName);
            return match.Success;
        }

        private static List<List<string>>[] separateUsersInSheets(List<List<string>> cells) {
            int length = cells.Count - 1;
            List<List<string>>[] sheets = new List<List<string>>[length];
            for (int i = 0; i < length; i++) {
                List<List<string>> tmp = new List<List<string>>();
                tmp.Add(cells[0]);
                tmp.Add(cells[i+1]);
                sheets[i] = tmp;
            }
            return sheets;
        }

        private static string readFile(string path) {
            return File.ReadAllText(path);
        }

        private static string[] readLines(string path) {
            String text = File.ReadAllText(path);
            List<int> indexOfLines = new List<int>();
            indexOfLines.Add(0);
            string[] lines = Regex.Split(text, "(\"\n\")");
            int effectiveLines = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i] != "\"\n\"")
                {
                    effectiveLines++;
                }
            }
            string[] realLines = new string[effectiveLines];
            int cpt = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i] != "\"\n\"" && cpt < realLines.Length)
                {
                    realLines[cpt] = lines[i];
                    cpt++;
                }
            }
            return realLines;
        }

        private static List<List<string>> getCells(string[] fileContent)
        {
            List<List<string>> cells = new List<List<string>>();
            foreach (string line in fileContent)
            {
                cells.Add(new List<string>(Regex.Split(line, "(\",\")")));
            }

            for (int i = 0; i < cells.Count; i++ )
            {
                List<string> tmpRow = new List<string>();
                foreach (string column in cells[i])
                {
                    if (column != "\",\"")
                    {
                        tmpRow.Add(column.TrimStart('"'));
                    }
                }
                cells[i] = tmpRow;
            }         
            return cells;
        }

        private static List<List<string>> cleanCells(List<List<string>> cells)
        {
            for (int j = 0; j < cells.Count; j++)
            {
                List<string> line = cells[j];
                for (int i = 0; i < line.Count; i++) {
                    if(line[i] == null){
                        line[i] = "";
                    }
                    line[i] = line[i].Replace("\n", "");
                    line[i] = line[i].Replace("\"", "");
                    if (line[i].Length > 0 && line[i][0] == ',') {
                        line[i] = line[i].Substring(1, line[i].Length -1);
                    }
                }
            }
            return cells;
        }

        private static List<List<string>> emptyCells(List<List<string>> cells)
        {
            for (int j = 0; j < cells.Count; j++)
            {
                List<string> line = cells[j];
                for (int i = 0; i < line.Count; i++)
                {
                    if(line[i] == "N/A" || line[i] == "Non"){
                        line[i] = "";
                    }
                }
            }
            return cells;
        }

        private static List<List<string>> removeCells(List<List<string>> cells)
        {
            List<int> indexesToDelete = new List<int>();
            for (int j = 0; j < cells.Count; j++)
            {
                List<string> line = cells[j];
                for (int i = 0; i < line.Count; i++)
                {
                    if (line[i] == "" && isNotInList(indexesToDelete, i))
                    {
                        indexesToDelete.Add(i);
                    }
                }
            }
            List<string> tmp0 = new List<string>();
            List<string> tmp1 = new List<string>();

            for (int j = 0; j < cells[0].Count; j++)
            {
                if(cells[0][j] != "" && cells[1][j] != "") {
                    tmp0.Add(cells[0][j]);
                    tmp1.Add(cells[1][j]);
                }
            }
            cells[0] = tmp0;
            cells[1] = tmp1;
            return cells;
        }

        private static bool isNotInList(List<int> indexList, int value) {
            foreach (int index in indexList)
            {
                if (index == value)
                    return false;
            }
            return true;
        }

        private static List<int> decreaseIndexes(List<int> indexList) {
            for (int i = 0; i < indexList.Count; i++) {
                if(indexList[i] != 0) {
                    indexList[i] -= 1;
                }
            }
            return indexList;
        }
    }
}
