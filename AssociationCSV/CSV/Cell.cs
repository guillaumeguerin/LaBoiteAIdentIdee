using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssociationCSV
{
    class Cell
    {
        private string cellContent;

        public Cell(string content)
        {
            this.setCell(content);
        }

        public string getCell()
        {
            return cellContent;
        }

        public void setCell(string content)
        {
            cellContent = content;
        }
    }
}
