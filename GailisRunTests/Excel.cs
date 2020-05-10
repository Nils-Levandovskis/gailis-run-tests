using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace GailisRunTests
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        private bool isOpen;
        Workbook wb;
        Worksheet ws;
        public Workbook Wb { get => wb; set => wb = value; }
        public Worksheet Ws { get => ws; set => ws = value; }
        public bool IsOpen { get => isOpen; }

        //public Excel()
        //{

        //}
        public Excel(string path, int sheet)
        {
            this.path = path ?? throw new ArgumentNullException(nameof(path));
            Wb = excel.Workbooks.Open(path);
            Ws = Wb.Worksheets[sheet];
            isOpen = true;
        }
        public void createNewWorkbook()
        {
            Wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Ws = Wb.Worksheets[1];

        }
        public void createNewSheet()
        {
            Wb.Worksheets.Add(After: Ws);
        }
        public void createNewSheet(bool select)
        {
            if (select)
            {
                Ws = Wb.Worksheets.Add(After: Ws);
            }
        }
        public int getSheetID()
        {
            return Ws.Index;
        }
        public void setSheetName(string name)
        {
            Ws.Name = name;
        }
        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (Ws.Cells[i, j].Value2 != null)
            {
                var cellValue = Ws.Cells[i, j].Value2;

                return (string)Ws.Cells[i, j].Text;
            }
            else return "";
        }
        public string ReadCell(Cell cell)
        {
            return ReadCell(cell.GetRow(), cell.GetCol());
        }
        public void ChangeSheet(int sheet, bool save)
        {
            if (save)
            {
                Save();
            }
            Ws = Wb.Worksheets[sheet];
        }
        public void AlterCell(int i, int j, string value)
        {
            i++;
            j++;
            Ws.Cells[i, j].Value2 = value;

        }
        public void AlterCell(Cell cell, string value)
        {
            AlterCell(cell.GetRow(), cell.GetCol(), value);
        }
        public void AlterCell(int i, int j, int value)
        {
            i++;
            j++;
            Ws.Cells[i, j].Value2 = value;
        }
        public void AlterCell(Cell cell, int value)
        {
            AlterCell(cell.GetRow(), cell.GetCol(), value);
        }
        public void Save()
        {
            Wb.Save();
        }
        public void Close()
        {
            isOpen = false;
            wb.Close(0);
            excel.Quit();
        }
        ~Excel()
        {
            if (isOpen)
            {
                Console.WriteLine(path);
                wb.Close(0);
                excel.Quit();
            }
        }
    }
}
