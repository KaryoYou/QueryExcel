using System.IO;
using System.Data;
using System.Data.OleDb;

namespace QueryExcel
{
    internal class Excel
    {
        public readonly string excelFilePath;
        public readonly bool isExcelFile = false;
        public readonly DataTable excelSheets;

        public Excel(string filePath)
        {
            if (IsExcelFileCheck(filePath))
            {
                excelFilePath = filePath;
                isExcelFile = true;
                excelSheets = ReadExcelSheets(string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties='Excel 12.0;'", filePath));
            }
        }

        private bool IsExcelFileCheck(string filePath)
        {
            string extension = Path.GetExtension(filePath);
            return extension == ".xls" || extension == ".xlsx";
        }

        private DataTable ReadExcelSheets(string connStr)
        {
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                conn.Open();
                DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                return dtSheetName;
            }
        }
    }
}
