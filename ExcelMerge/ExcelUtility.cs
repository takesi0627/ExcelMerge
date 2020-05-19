using System;
using System.IO;
using ClosedXML.Excel;

// using NPOI.SS.UserModel;
// using NPOI.HSSF.UserModel;
// using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public class ExcelUtility
    {

        // public static void RemoveEmptyRows(IXLWorksheet sheet)
        // {
        //     bool isRowEmpty = false;
        //     for (int i = 0; i <= sheet.LastRowNum; i++)
        //     {
        //         if (sheet.GetRow(i) == null)
        //         {
        //             isRowEmpty = true;
        //             sheet.ShiftRows(i + 1, sheet.LastRowNum + 1, -1);
        //             i--;
        //             continue;
        //         }
        //         for (int j = 0; j < sheet.GetRow(i).LastCellNum; j++)
        //         {
        //             if (sheet.GetRow(i).GetCell(j) == null || GetCellStringValue(sheet.GetRow(i).GetCell(j)) == string.Empty)
        //             {
        //                 isRowEmpty = true;
        //             }
        //             else
        //             {
        //                 isRowEmpty = false;
        //                 break;
        //             }
        //         }
        //         if (isRowEmpty == true)
        //         {
        //             sheet.ShiftRows(i + 1, sheet.LastRowNum + 1, -1);
        //             i--;
        //         }
        //     }
        // }
        // public static object GetCellValue(ICell cell)
        // {
        //     if (cell == null)
        //         return null;
        //
        //     return GetCellValue(cell, cell.CellType);
        // }

        // private static object GetCellValue(ICell cell, CellType type)
        // {
        //     if (cell != null)
        //     {
        //         switch (type)
        //         {
        //             case CellType.Numeric:
        //                 if (DateUtil.IsCellDateFormatted(cell))
        //                 {
        //                     return cell.DateCellValue;
        //                 }
        //                 else
        //                 {
        //                     return cell.NumericCellValue;
        //                 }
        //             case CellType.String:
        //                 return cell.StringCellValue;
        //             case CellType.Boolean:
        //                 return cell.BooleanCellValue;
        //             case CellType.Formula:
        //                 return GetCellValue(cell, cell.CachedFormulaResultType);
        //         }
        //     }
        //
        //     return string.Empty;
        // }

        // public static string GetCellStringValue(ICell cell)
        // {
        //     if (cell == null)
        //         return string.Empty;
        //
        //     return GetCellValue(cell).ToString();
        // }

        public static void CreateWorkbook(string path, ExcelWorkbookType workbookType)
        {
            if (!ValidateExtension(path, workbookType))
                throw new ArgumentException("The specified Excel type and path extension do not match.");

            var workbook = CreateWorkbook(workbookType);
            // var sheet = workbook.CreateSheet();

            workbook.SaveAs(path);

            // using (var fileStream = new FileStream(path, FileMode.Create))
            // {
            //     workbook.Write(fileStream);
            // }
        }
        private static IXLWorkbook CreateWorkbook(ExcelWorkbookType workbookType)
        {
            switch (workbookType)
            {
                // case ExcelWorkbookType.XLS: return new HSSFWorkbook() as IWorkbook;
                case ExcelWorkbookType.XLSX: return new XLWorkbook();
                default: break;
            }

            throw new ArgumentException("The specified excel type is not supported instantiating.");
        }

        private static bool ValidateExtension(string path, ExcelWorkbookType workbookType)
        {
            switch (workbookType)
            {
                case ExcelWorkbookType.XLS: return Path.GetExtension(path) == ".xls";
                case ExcelWorkbookType.XLSX: return Path.GetExtension(path) == ".xlsx";
                default: break;
            }

            return false;
        }

        public static ExcelWorkbookType GetWorkbookType(string path)
        {
            var extension = Path.GetExtension(path);
            switch (extension)
            {
                case ".xls": return ExcelWorkbookType.XLS;
                case ".xlsx": return ExcelWorkbookType.XLSX;
                default: break;
            }

            return ExcelWorkbookType.None;
        }

        // public static ExcelWorkbookType GetWorkboolTypeStrict(string path)
        // {
        //     var type = GetWorkbookType(path);
        //
        //     if (type == ExcelWorkbookType.None)
        //     {
        //         // if (IsXLS(path))
        //         //     type = ExcelWorkbookType.XLS;
        //         if (IsXLSX(path))
        //             type = ExcelWorkbookType.XLSX;
        //     }
        //
        //     return type;
        // }

        // public static bool IsXLS(string path)
        // {
        //     try
        //     {
        //         return WorkbookFactory.Create(path) is HSSFWorkbook;
        //     }
        //     catch
        //     {
        //         return false;
        //     }
        // }

        // public static bool IsXLSX(string path)
        // {
        //     try
        //     {
        //         return WorkbookFactory.Create(path) is XSSFWorkbook;
        //     }
        //     catch
        //     {
        //         return false;
        //     }
        // }
    }
}
