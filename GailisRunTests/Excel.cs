using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
/*
 Basic excel workbook manipulation functionality
     */
namespace GailisRunTests
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        private bool isOpen;
        public Workbook Wb { get; set; }
        public Worksheet Ws { get; set; }
        public bool IsOpen { get => isOpen; }

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
        public bool SaveAs(string fileName)
        {
            try
            {
                string filePath = Path.GetFullPath(fileName);
                Wb.SaveCopyAs(filePath);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0}: {1}", e.Message, e.ToString());
                return false;
            }
            return true;

        }
        public void Close()
        {
            isOpen = false;
            Wb.Close(0);
            excel.Quit();
            Marshal.FinalReleaseComObject(Ws);
            Marshal.FinalReleaseComObject(Wb);
            Marshal.FinalReleaseComObject(excel);
            Console.WriteLine("Closed {0} via Close()", path);
        }
        ~Excel()
        {
            if (isOpen)
            {
                Wb.Close(0);
                excel.Quit();
                Marshal.FinalReleaseComObject(Ws);
                Marshal.FinalReleaseComObject(Wb);
                Marshal.FinalReleaseComObject(excel);
                Console.WriteLine("Closed {0} via destructor", path);
            }
        }
    }
}
