using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AssociationCSV
{
    class Row
    {
        private List<Column> rows;

        public Row() {
            rows = new List<Column>();
        }

        public Row(List<Column> myRows) {
            rows = myRows;
        }

        public List<Column> getRows() {
            return rows;
        }

        public void setRows(List<Column> columnsList) {
            rows = columnsList;
        }

        public void setRows(string columnsList)
        {
            string separator = "\"\n\"";
            string[] lines = Regex.Split(columnsList, separator);
            List<Column> result = new List<Column>();

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i] != separator)
                {
                    Column tmp = new Column();
                    tmp.setCells(lines[i]);
                    result.Add(tmp);
                }
            }

            this.setRows(result);
        }

    }
}
