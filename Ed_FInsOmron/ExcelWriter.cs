using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace Ed_FInsOmron.Excel
{
    public interface IExcelManager
    {
        void WriteData(string sheetName, int row, int col,int numCol, string[] value);
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

        public void WriteData(string sheetName, int row, int col, int numCol, string[] value)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            ICell cell;

            IRow excelRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
            for(int i = col, j = 0; i < numCol; i++, j++)
            {
                cell = excelRow.GetCell(i) ?? excelRow.CreateCell(i);
                cell.SetCellValue(value[j]);
            }

            using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(file);
            }
        }
    }
}