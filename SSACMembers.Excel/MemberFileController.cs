using System.IO;
using NPOI.HSSF.Record;
using NPOI.XSSF.UserModel;

namespace SSACMembers.Excel
{
    public class MemberFileController
    {
        public static void TransformMemberFile()
        {
            ExcelFileReader reader = new ExcelFileReader(@"C:\Users\mercu\Downloads\SSAC Member List.xlsx");
            ExcelFileReader writer = new ExcelFileReader(@"C:\Users\mercu\Downloads\Cell Group Members.xlsx");

            var readWorksheet = reader.GetWorksheet("Cell 1");
            var writeWorksheet = reader.GetWorksheet("Sheet1");

            for (int row = 2; row <= readWorksheet.LastRowNum; row++)
            {
                for (int column = 0; column < 30; column++)
                {
                    if (string.IsNullOrEmpty(readWorksheet.GetRow(row).GetCell(column).StringCellValue))
                        break;
                    
                    string cellValue = readWorksheet.GetRow(row).GetCell(column).StringCellValue.Trim();

                    readWorksheet.GetRow(2).GetCell(0).SetCellValue("Jame");

                    //using (var file2 = new FileStream((@"C:\Users\mercu\Downloads\Cell Group Members.xlsx", FileMode.Create, FileAccess.ReadWrite))
                    //{
                    //    //readWorksheet.Write(file2);
                    //    file2.Close();
                    //}

                    // https://social.msdn.microsoft.com/Forums/vstudio/en-US/d1c5e191-135b-45c0-9f88-cc3e02849257/npoi-how-to-write-to-an-xlsx-excel-file?forum=csharpgeneral
                }
            }
        }

        public static void TransformMemberFileNew()
        {
            XSSFWorkbook wb1 = null;
            using (var file = new FileStream(@"C:\Users\mercu\Downloads\Cell Group Members.xlsx", FileMode.Open, FileAccess.ReadWrite))
            {
                wb1 = new XSSFWorkbook(file);
            }
            wb1.GetSheetAt(0).GetRow(0).GetCell(0).SetCellValue("Sample");

            using (var file2 = new FileStream(@"C:\Users\mercu\Downloads\SSAC Member List.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
                wb1.Write(file2);
                file2.Close();
            }
            var file = new NPOIExcel
        }
    }
}
