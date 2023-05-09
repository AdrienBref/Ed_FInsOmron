using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace Ed_FInsOmron.Excel
{
    public interface IExcelManager
    {
        void WriteData(int row, int col, string value);
    }

    public class ExcelManager : IExcelManager
    {
        private readonly XSSFWorkbook _workbook;
        private string fileName = "";

        public ExcelManager(string fileName)
        {
            this.fileName = fileName;

            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                _workbook = new XSSFWorkbook(file);
            }
        }


        /// <summary>
        /// Set Cell Value with the Row and col cordenates
        /// </summary>
        /// <param name="row">integer that give the value of the row.</param>
        /// <param name="col">integer that give the value of the row</param>
        /// <param name="value">String with the value towrite</param>
        public void WriteData(int row, int col, string value)
        {
            ISheet sheet = _workbook.GetSheetAt(0);
            ICell cell;

            IRow excelRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
            cell = excelRow.GetCell(col) ?? excelRow.CreateCell(col);
            cell.SetCellValue(value);
            

            using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(file);
            }
        }
    }
}