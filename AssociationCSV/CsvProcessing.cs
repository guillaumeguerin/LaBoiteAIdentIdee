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
                    /*
                    string[] fileContent = readLines(path);
                    List<List<string>> cells = getCells(fileContent);
                    List<List<string>>[] sheets = separateUsersInSheets(cells);
                     
                    foreach (List<List<string>> sheet in sheets)
                    {
                        processSheet(sheet);
                        HTMLRendering.generatePDFs(sheet);
                    }
                    */

                    string[] fileContent = readLines(path);
                    List<List<string>> cells = getCells(fileContent);
                    List<List<string>>[] sheets = separateUsersInSheets(cells);
                    /*string fileContent = readFile(path);
                    Row sheet = convertToCells(fileContent);
                    List<Row> suscriberSheets = separateUsersInSheets(sheet);
                    foreach (Row mySheet in suscriberSheets)*/
                    foreach (List<List<string>> mySheet in sheets)
                    {
                        processSheet(mySheet);
                        string name = HTMLRendering.getLastName(mySheet).ToUpper();
                        string firstname = HTMLRendering.getFirstName(mySheet).ToUpper();
                        string fileName = "CV_" + firstname + name;
                        string text = HTMLRendering.buildHtml(mySheet);
                        Form2 previewForm = new Form2(text, name, firstname);
                        DialogResult dr = previewForm.ShowDialog();
                        if(dr == DialogResult.OK) {
                            text = Form2.text;
                        }
                        string path2 = Directory.GetCurrentDirectory() + "\\" + fileName + ".html";
                        File.WriteAllText(Directory.GetCurrentDirectory() + "\\" + fileName + ".html", text, Encoding.GetEncoding("windows-1252"));

                        string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        Directory.SetCurrentDirectory(exeDir);
                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                        startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                        startInfo.FileName = "cmd.exe";
                        startInfo.Arguments = "/C phantomjs pdf.js " + fileName + ".html " + fileName + ".pdf";
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


        private static Row processSheet(Row mySheet)
        {
            mySheet = cleanCells(mySheet);
            mySheet = emptyCells(mySheet);
            mySheet = removeCells(mySheet);

            return mySheet;
        }

        private static Row removeCells(Row mySheet)
        {
            List<int> indexesToDelete = new List<int>();
            List<Column> rows = mySheet.getRows();
            for (int j = 0; j < rows.Count; j++)
            {
                List<Cell> cells = rows[j].getCells();
                for (int i = 0; i < cells.Count; i++)
                {
                    if (cells[i].getCell() == "" && isNotInList(indexesToDelete, i))
                    {
                        indexesToDelete.Add(i);
                    }
                }
            }
            List<Cell> tmp0 = new List<Cell>();
            List<Cell> tmp1 = new List<Cell>();

            for (int j = 0; j < rows[0].getCells().Count; j++)
            {
                if (rows[0].getCells()[j].getCell() != "" && rows[1].getCells()[j].getCell() != "")
                {
                    tmp0.Add(new Cell(rows[0].getCells()[j].getCell()));
                    tmp1.Add(new Cell(rows[1].getCells()[j].getCell()));
                }
            }
            rows[0].setCells(tmp0);
            rows[1].setCells(tmp1);
            mySheet.setRows(rows);
            return mySheet;
        }

        private static Row emptyCells(Row mySheet)
        {
            List<Column> rows = mySheet.getRows();
            for (int j = 0; j < rows.Count; j++)
            {
                List<Cell> line = rows[j].getCells();
                for (int i = 0; i < line.Count; i++)
                {
                    if (line[i].getCell() == "N/A" || line[i].getCell() == "Non")
                    {
                        line[i].setCell("");
                    }
                }
                rows[j].setCells(line);
            }
            mySheet.setRows(rows);
            return mySheet;
        }

        private static Row cleanCells(Row mySheet)
        {
            List<Column> rows = mySheet.getRows();
            for (int j = 0; j < rows.Count; j++)
            {
                Column line = rows[j];
                List<Cell> cells = line.getCells();
                for (int i = 0; i < cells.Count; i++)
                {
                    if (cells[i] == null)
                    {
                        cells[i].setCell("");
                    }
                    cells[i].setCell(cells[i].getCell().Replace("\n", ""));
                    cells[i].setCell(cells[i].getCell().Replace("\"", ""));
                    if (cells[i].getCell().Length > 0 && cells[i].getCell()[0] == ',')
                    {
                        cells[i].setCell(cells[i].getCell().Substring(1, cells[i].getCell().Length - 1));
                    }
                }
                line.setCells(cells);
            }
            mySheet.setRows(rows);
            return mySheet;
        }

        private static List<Row> separateUsersInSheets(Row sheet)
        {
            List<Row> sheets = new List<Row>();
            for (int i = 1; i < sheet.getRows().Count; i++)
            {
                Row tmp = new Row();
                tmp.getRows().Add(sheet.getRows()[0]);
                tmp.getRows().Add(sheet.getRows()[i]);
                sheets.Add(tmp);
            }
            return sheets;
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

        private static Row convertToCells(string fileContent)
        {
            Row result = new Row();
            result.setRows(fileContent);
            return result;
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
