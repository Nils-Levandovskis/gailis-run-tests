using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/*
 X,Y coordinate class to interface with excel cells
*/
namespace GailisRunTests
{
    class Cell
    {
        private int row;
        private int col;
        public Cell(int i, int j)
        {
            this.row = i + 1;
            this.col = j + 1;
        }
        public void ZeroIndexSet(int i, int j)
        {
            this.row = i + 1;
            this.col = j + 1;
        }
        public void ActualIndexSet(int i, int j)
        {
            this.row = i;
            this.col = j;
        }
        
        public int GetRow()
        {
            return row - 1;
        }
        public int GetRow(bool actual)
        {
            if (actual) return row;
            return row - 1;
        }
        public int GetCol()
        {
            return col - 1;
        }
        public int GetCol(bool actual)
        {
            if (actual) return col;
            return col - 1;
        }
        public void Down()
        {
            this.row++;
        }
        public void Right()
        {
            this.col++;
        }


    }
}
