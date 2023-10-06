using OpenBomWpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpenBomWpf
{
    public class Range
    {

        //==============================================
        // CONSTRUCTORS
        //==============================================
        public Range(Cell topCell, Cell bottomCell)
        {
            this.TopCell = topCell;
            this.BottomCell = bottomCell;
        }

        public Range(Cell singleCell)
        {
            this.TopCell = singleCell;
            this.BottomCell = singleCell;
        }


        //==============================================
        // PROPERTIES
        //==============================================
        public Cell TopCell { get; private set; }
        public Cell BottomCell { get; private set; }

        public int[] RowIndices
        {
            get
            {
                List<int> list = new List<int>();

                for (int index = TopCell.RowIndex; index <= BottomCell.RowIndex; index++) list.Add(index);

                return list.ToArray();
            }
        }

        public int[] ColumnIndices
        {
            get
            {
                List<int> list = new List<int>();

                for (int index = TopCell.ColumnIndex; index <= BottomCell.ColumnIndex; index++) list.Add(index);

                return list.ToArray();
            }
        }


        //==============================================
        // METHODS
        //==============================================
        public Excel.Range GetExcelRange(Excel.Worksheet sheet)
        {
            Excel.Range topCell = (Excel.Range)sheet.Cells[this.TopCell.RowIndex, this.TopCell.ColumnIndex];
            Excel.Range bottomCell = (Excel.Range)sheet.Cells[this.BottomCell.RowIndex, this.BottomCell.ColumnIndex];
            return sheet.Application.get_Range(topCell, bottomCell);
        }


        public Excel.Range GetExcelRange(Excel.Range range)
        {
            Excel.Range topCell = (Excel.Range)range.Cells[this.TopCell.RowIndex, this.TopCell.ColumnIndex];
            Excel.Range bottomCell = (Excel.Range)range.Cells[this.BottomCell.RowIndex, this.BottomCell.ColumnIndex];
            return range.get_Range(topCell, bottomCell);
        }
    }
}
