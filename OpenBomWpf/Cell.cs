using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenBomWpf
{
    public class Cell
    {

        //==============================================
        // CONSTRUCTORS
        //==============================================
        public Cell(int oneBasedRow, int oneBasedColumn)
        {
            if (oneBasedColumn < 1 || oneBasedRow < 1)
            {
                throw new ArgumentException("Row and column indexes are one-based.");
            }

            this.RowIndex = oneBasedRow;
            this.ColumnIndex = oneBasedColumn;
        }


        //==============================================
        // PROPERTIES
        //==============================================
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }


        //==============================================
        // METHODS
        //==============================================
        public static Cell Add(Cell c, int rows, int columns)
        {
            return new Cell(c.RowIndex + rows, c.ColumnIndex + columns);
        }
    }
}

