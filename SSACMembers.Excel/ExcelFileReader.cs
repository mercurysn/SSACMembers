using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SSACMembers.Excel
{
    public class ExcelFileReader
    {
        private readonly string _fileName;

        public ExcelFileReader(string fileName)
        {
            _fileName = fileName;
        }

        public ISheet GetWorksheet(string sheetName)
        {
            var fileExt = Path.GetExtension(_fileName);

            //Declare the sheet interface
            ISheet sheet;

            //Get the Excel file according to the extension
            if (fileExt != null && fileExt.ToLower() == ".xls")
            {
                //Use the NPOI Excel xls object
                HSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(_fileName, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new HSSFWorkbook(file);
                }

                //Assign the sheet
                sheet = hssfwb.GetSheet(sheetName);
            }
            else //.xlsx extension
            {
                //Use the NPOI Excel xlsx object
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(_fileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    hssfwb = new XSSFWorkbook(file);
                }

                //Assign the sheet
                sheet = hssfwb.GetSheet(sheetName);
            }

            return sheet;
        }
    }
}
