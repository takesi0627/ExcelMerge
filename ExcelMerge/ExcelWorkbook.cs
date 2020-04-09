using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }

        public List<string> SheetNames { get; private set; }

        private string rawFilePath;

        private XSSFWorkbook rawWorkbook;

        private string tmpFileName;

        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
            SheetNames = new List<string>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);

            string tmpFile = Path.GetTempFileName();
            File.Copy(path, tmpFile, true);

            var srcWb = new XSSFWorkbook(tmpFile);
            var wb = new ExcelWorkbook();

            wb.rawFilePath = path;
            wb.rawWorkbook = srcWb;

            for (int i = 0; i < srcWb.NumberOfSheets; i++)
            {
                var srcSheet = srcWb.GetSheetAt(i);
                wb.Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
                wb.SheetNames.Add(srcSheet.SheetName);
            }

            return wb;
        }

        public void DumpByCreate()
        {
            // 尝试直接通过封装的数据来创建出原始 excel 表，比较冒险，可能有缺少的内容导致写入的 excel 错误
            // 但是这样写入成功之后，界面操作修改都会变得简单
            var wb = new XSSFWorkbook();

            foreach (var sheetName in SheetNames)
            {
                var table = wb.CreateSheet(sheetName);

                var sheetWrap = Sheets[sheetName];

                foreach (var rowWrap in sheetWrap.Rows)
                {
                    var row = table.CreateRow(rowWrap.Key);

                    // if (row is null)
                    // {
                    //     break;
                    // }

                    var i = 0;
                    foreach (var cellWrap in rowWrap.Value.Cells)
                    {
                        var cell = row.CreateCell(i);
                        // var cell = row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        XSSFCellStyle cellstyle = wb.CreateCellStyle() as XSSFCellStyle;
                        if (cellWrap.OriginalCellStyle != null)
                        {
                            cellstyle.CloneStyleFrom(cellWrap.OriginalCellStyle);
                        }
                        
                        cell.CellStyle = cellstyle;
                        cell.SetCellValue(cellWrap.Value);
                        i++;
                    }
                }
            }

            using (FileStream stream = new FileStream(@"D:/test.xlsx", FileMode.Create, FileAccess.Write))
            {
                wb.Write(stream);
            }
        }

        public void Dump()
        {
            var workbook = rawWorkbook;

            foreach (var sheetName in SheetNames)
            {
                var table = workbook.GetSheet(sheetName);

                var sheetWrap = Sheets[sheetName];

                foreach (var rowWrap in sheetWrap.Rows)
                {
                    var row = table.GetRow(rowWrap.Key);

                    if (row is null)
                    {
                        break;
                    }

                    var i = 0;
                    foreach (var cellWrap in rowWrap.Value.Cells)
                    {
                        var cell = row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        cell.SetCellValue(cellWrap.Value);
                        i++;
                    }
                }
            }

            using (FileStream stream = new FileStream(rawFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }

            // using (var fs = File.OpenWrite(@"D:/test.xlsx"))
            // {
            //     workbook.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。
            //     Console.WriteLine("生成成功");
            // }
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            if (Path.GetExtension(path) == ".csv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else if (Path.GetExtension(path) == ".tsv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else
            {
                var wb = WorkbookFactory.Create(path);
                for (int i = 0; i < wb.NumberOfSheets; i++)
                    yield return wb.GetSheetAt(i).SheetName;
            }
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromCsv(path, config));

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromTsv(path, config));

            return wb;
        }
    }
}
