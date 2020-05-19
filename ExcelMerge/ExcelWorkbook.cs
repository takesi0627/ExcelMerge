using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NetDiff;
// using NPOI.HSSF.UserModel;
// using NPOI.SS.UserModel;
// using NPOI.XSSF.UserModel;
using System.Diagnostics;
using ClosedXML.Excel;
using NPOI.OpenXmlFormats.Dml.WordProcessing;


namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }

        public List<string> SheetNames { get; private set; }

        private string rawFilePath;

        private IXLWorkbook rawWorkbook;

        private string tmpFileName;

        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
            SheetNames = new List<string>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config, string sheetName)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);

            string tmpFile = Path.GetTempFileName() + ".xlsx";
            File.Copy(path, tmpFile, true);

            var rawWorkbook = new XLWorkbook(tmpFile);
            var wb = new ExcelWorkbook();

            wb.rawFilePath = path;
            wb.rawWorkbook = rawWorkbook;

            rawWorkbook.TryGetWorksheet(sheetName, out var srcSheet);

            // foreach (var srcSheet in rawWorkbook.Worksheets)
            // {
            wb.Sheets.Add(srcSheet.Name, ExcelSheet.Create(srcSheet, config));
            wb.SheetNames.Add(srcSheet.Name);
            // }

            return wb;
        }

        // public void DumpByCreate()
        // {
        //     // 尝试直接通过封装的数据来创建出原始 excel 表，比较冒险，可能有缺少的内容导致写入的 excel 错误
        //     // 但是这样写入成功之后，界面操作修改都会变得简单
        //     var wb = new XSSFWorkbook();
        //
        //     foreach (var sheetName in SheetNames)
        //     {
        //         var table = wb.CreateSheet(sheetName);
        //
        //         var sheetWrap = Sheets[sheetName];
        //
        //         foreach (var rowWrap in sheetWrap.Rows)
        //         {
        //             var row = table.CreateRow(rowWrap.Key);
        //
        //             // if (row is null)
        //             // {
        //             //     break;
        //             // }
        //
        //             var i = 0;
        //             foreach (var cellWrap in rowWrap.Value.Cells)
        //             {
        //                 var cell = row.CreateCell(i);
        //                 // var cell = row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        //
        //                 XSSFCellStyle cellstyle = wb.CreateCellStyle() as XSSFCellStyle;
        //                 if (cellWrap.OriginalCellStyle != null)
        //                 {
        //                     cellstyle.CloneStyleFrom(cellWrap.OriginalCellStyle);
        //                 }
        //                 
        //                 cell.CellStyle = cellstyle;
        //                 cell.SetCellValue(cellWrap.Value);
        //                 i++;
        //             }
        //         }
        //     }
        //
        //     using (FileStream stream = new FileStream(@"D:/test.xlsx", FileMode.Create, FileAccess.Write))
        //     {
        //         wb.Write(stream);
        //     }
        // }

        public void Dump(string sheetName, ExcelSheetDiff sheetDiff, bool isLeft)
        {
            var workbook = rawWorkbook;

            workbook.TryGetWorksheet(sheetName, out var rawSheet);

            var tableModified = false;

            // 添加需要的行，让原始表格和diff行数保持一致
            foreach (KeyValuePair<int, ExcelRowDiff> sheetDiffRow in sheetDiff.Rows)
            {
                if (sheetDiffRow.Value.IsRemoved() && !isLeft)
                {
                    Debug.Print("ShiftAddRight : " + sheetDiffRow.ToString());
                    // rawSheet.ShiftRows(sheetDiffRow.Key, rawSheet.LastRowNum, 1);
                    rawSheet.Row(sheetDiffRow.Key + 1).InsertRowsAbove(1);
                    // if (sheetDiffRow.Key <= rawSheet.RowCount())
                    // {
                    //     Debug.Print("ShiftAddRight : " + sheetDiffRow.ToString());
                    //     // rawSheet.ShiftRows(sheetDiffRow.Key, rawSheet.LastRowNum, 1);
                    //     rawSheet.Row(sheetDiffRow.Key + 1).InsertRowsAbove(1);
                    // }
                    
                    // rawSheet.CreateRow(sheetDiffRow.Key);
                }

                if (sheetDiffRow.Value.IsAdded() && isLeft)
                {
                    Debug.Print("ShiftAddLeft : " + sheetDiffRow.ToString());
                    rawSheet.Row(sheetDiffRow.Key + 1).InsertRowsAbove(1);
                    // if (sheetDiffRow.Key <= rawSheet.RowCount())
                    // {
                    //     Debug.Print("ShiftAddLeft : " + sheetDiffRow.ToString());
                    //     rawSheet.Row(sheetDiffRow.Key).InsertRowsAbove(1);
                    //     // rawSheet.ShiftRows(sheetDiffRow.Key, rawSheet.LastRowNum, 1);
                    // }
                    // rawSheet.CreateRow(sheetDiffRow.Key);
                }
            }

            // 逐行比对修改
            foreach (var rowDiff in sheetDiff.Rows)
            {
                if (rowDiff.Value.LeftEmpty())
                {
                    continue;
                }

                if (rowDiff.Value.RightEmpty())
                {
                    continue;
                }

                if (!rowDiff.Value.NeedMerge())
                {
                    continue;
                }

                // rawSheet.Row(rowDiff.Key);
                var rawRow = rawSheet.Row(rowDiff.Key + 1);

                foreach (var cellDiff in rowDiff.Value.Cells)
                {
                    var rawCol = cellDiff.Key + 1;
                    Debug.Print(rawCol.ToString());
                    var rawCell = rawRow.Cell(rawCol);
                    // var rawCell = rawRow.GetCell(cellDiff.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    ExcelCell targetWrap = null;
                    if (isLeft)
                    {
                        if (cellDiff.Value.MergeStatus == ExcelCellMergeStatus.UseRight)
                        {
                            targetWrap = cellDiff.Value.DstCell;
                        }
                    }
                    else
                    {
                        
                        if (cellDiff.Value.MergeStatus == ExcelCellMergeStatus.UseLeft)
                        {
                            targetWrap = cellDiff.Value.SrcCell;
                        }
                            
                    }

                    if (targetWrap != null)
                    {
                        targetWrap.RawCell?.CopyTo(rawCell);
                        // var style = workbook.CreateCellStyle();
                        //
                        // style.CloneStyleFrom(targetWrap.RawCell.CellStyle);
                        // rawCell.CellStyle = style;
                        // rawCell.SetCellType(targetWrap.RawCell.CellType);
                        //
                        //
                        // switch (targetWrap.RawCell.CellType)
                        // {
                        //     case CellType.Unknown:
                        //         break;
                        //     case CellType.Numeric:
                        //         rawCell.SetCellValue(targetWrap.RawCell.NumericCellValue);
                        //         break;
                        //     case CellType.String:
                        //         rawCell.SetCellValue(targetWrap.RawCell.StringCellValue);
                        //         break;
                        //     case CellType.Formula:
                        //         rawCell.SetCellValue(targetWrap.RawCell.CellFormula);
                        //         break;
                        //     case CellType.Blank:
                        //         break;
                        //     case CellType.Boolean:
                        //         rawCell.SetCellValue(targetWrap.RawCell.BooleanCellValue);
                        //         break;
                        //     case CellType.Error:
                        //         break;
                        //     default:
                        //         rawCell.SetCellValue(targetWrap.Value);
                        //         break;
                        // }

                        tableModified = true;
                    }
                    
                }
            }

            int index = 0;
            if (isLeft)
            {
                foreach (var rowDiff in sheetDiff.Rows)
                {
                    if (rowDiff.Value.LeftEmpty())
                    {
                        Debug.Print("ShiftBackLeft: " + rowDiff.ToString());
                        rawSheet.Row(index).Delete();
                    }
                    else
                    {
                        index++;
                    }

                }
            }
            else
            {
                foreach (var rowDiff in sheetDiff.Rows)
                {

                    if (rowDiff.Value.RightEmpty())
                    {
                        Debug.Print("ShiftBackRight: " +  rowDiff.ToString());
                        rawSheet.Row(index).Delete();
                        // if (index + 1 < rawSheet.RowCount())
                        // {
                        //     rawSheet.Row(index).Delete();
                        // }
                        
                    }
                    else
                    {
                        index++;
                    }
                }
            }

            if (tableModified)
            {
                workbook.SaveAs(rawFilePath);
                // using (FileStream stream = new FileStream(rawFilePath, FileMode.Create, FileAccess.Write))
                // {
                //     workbook.Write(stream);
                // }
            }
            

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
                var wb = new XLWorkbook(path);
                foreach (var wbWorksheet in wb.Worksheets)
                {
                    yield return wbWorksheet.Name;
                }
                // for (int i = 0; i < wb.NumberOfSheets; i++)
                //     yield return wb.GetSheetAt(i).SheetName;
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
