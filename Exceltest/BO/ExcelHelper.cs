using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceltest.BO
{
    class ExcelHelper: IDisposable
    {
        private string _filePath;
        private Excel.Worksheet _worksheet;
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private Excel.Range _range;

        //Excel.Application _excel = new Excel.Application();
        //Excel.Workbook xlWorkbook = xlApp.Workbooks();
        //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //Excel.Range xlRange = xlWorksheet.UsedRange;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close(null, null, null);
     
            }
            catch(Exception ex) {
            _workbook.Close(null, null, null); Console.WriteLine(ex.Message); }
        }

        internal object Get(string column, int row)
        {
            try
            {
                return _range.Cells[row, column].Value2.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
                return "empty";
        }

        internal bool Set(string column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                    _worksheet = _workbook.Sheets[1];
                    _range = _worksheet.UsedRange;

            }
                else
                {
                    //_workbook = _excel.Workbooks.Add();
                    //_filePath = filePath;
                    Console.WriteLine("File not found");
                    return false;
            }
            return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }
    }
}
