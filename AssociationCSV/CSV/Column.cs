using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AssociationCSV
{
    class Column
    {
        private List<Cell> cells;

        public Column() {
            cells = new List<Cell>();
        }

        public Column(List<Cell> myCells) {
            cells = myCells;
        }

        public List<Cell> getCells()
        {
            return cells;
        }

        public void setCells(List<Cell> cellList){
            cells = cellList;
        }

        public void setCells(string cellList)
        {
            string separator = "\",\"";
            string[] cells = Regex.Split(cellList, separator);
            List<Cell> result = new List<Cell>();
            for (int i = 0; i < cells.Length; i++)
            {
                if (cells[i] != separator)
                {
                    result.Add(new Cell(cells[i]));
                }
            }
            this.setCells(result);
        }
    }
}
