using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;

namespace ExcelMerge
{
    public abstract class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; } = new Dictionary<string, ExcelSheet>();

        public List<string> SheetNames { get; private set; } = new List<string>();

        protected string rawFilePath;


        public ExcelWorkbook(string path)
        {
            rawFilePath = path;
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            var ext = Path.GetExtension(path);
            if (ext == ".csv" || ext == ".tsv")
            {
                return new SVWorkbook(path, config);
            }
            else if (ext == ".xlsx" || ext == ".xlsm")
            {
                return new XLSWorkbook(path, config);
            }
            else
            {
                Debug.Assert(false, $"Invalid file type: {ext}");
                return null;
            }
        }

        public abstract void Dump(string sheetName, ExcelSheetDiff sheetDiff, bool isLeft);

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
    }

    internal class SVWorkbook : ExcelWorkbook
    {
        IList<IList<string>> values;

        internal SVWorkbook(string path, ExcelSheetReadConfig config) : base (path)
        {
            var extension = Path.GetExtension(path);
            if (extension == ".csv")
                Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromCsv(path, config));
            else if (extension == ".tsv")
                Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromTsv(path, config));
            else
            {
                Debug.Assert(false, $"Invalid file type: {extension}");
            }
        }

        public override void Dump(string sheetName, ExcelSheetDiff sheetDiff, bool isLeft)
        {

        }
    }

    internal class XLSWorkbook : ExcelWorkbook
    {
        private XSSFWorkbook rawWorkbook;

        internal XLSWorkbook(string path, ExcelSheetReadConfig config) : base (path)
        {
            string tmpFile = Path.GetTempFileName();
            File.Copy(path, tmpFile, true);

            rawWorkbook = new XSSFWorkbook(tmpFile);

            for (int i = 0; i < rawWorkbook.NumberOfSheets; i++)
            {
                var srcSheet = rawWorkbook.GetSheetAt(i);
                Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
                SheetNames.Add(srcSheet.SheetName);
            }
        }

        public override void Dump(string sheetName, ExcelSheetDiff sheetDiff, bool isLeft)
        {
            var workbook = rawWorkbook;

            var table = workbook.GetSheet(sheetName);

            var tableModified = false;

            // 添加需要的行，让原始表格和diff行数保持一致
            foreach (KeyValuePair<int, ExcelRowDiff> sheetDiffRow in sheetDiff.Rows)
            {
                if (sheetDiffRow.Value.IsRemoved() && !isLeft)
                {
                    if (sheetDiffRow.Key <= table.LastRowNum)
                    {
                        Debug.Print("ShiftAddRight : " + sheetDiffRow.ToString());
                        table.ShiftRows(sheetDiffRow.Key, table.LastRowNum, 1);
                    }

                    table.CreateRow(sheetDiffRow.Key);
                }

                if (sheetDiffRow.Value.IsAdded() && isLeft)
                {
                    if (sheetDiffRow.Key <= table.LastRowNum)
                    {
                        Debug.Print("ShiftAddLeft : " + sheetDiffRow.ToString());
                        table.ShiftRows(sheetDiffRow.Key, table.LastRowNum, 1);
                    }
                    table.CreateRow(sheetDiffRow.Key);
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

                var rawRow = table.GetRow(rowDiff.Key) ?? table.CreateRow(rowDiff.Key);

                foreach (var cellDiff in rowDiff.Value.Cells)
                {
                    var rawCell = rawRow.GetCell(cellDiff.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);
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
                        if (targetWrap.RawCell != null)
                        {
                            var style = workbook.CreateCellStyle();

                            style.CloneStyleFrom(targetWrap.RawCell.CellStyle);
                            rawCell.CellStyle = style;
                            rawCell.SetCellType(targetWrap.RawCell.CellType);

                            switch (targetWrap.RawCell.CellType)
                            {
                                case CellType.Unknown:
                                    break;
                                case CellType.Numeric:
                                    rawCell.SetCellValue(targetWrap.RawCell.NumericCellValue);
                                    break;
                                case CellType.String:
                                    rawCell.SetCellValue(targetWrap.RawCell.StringCellValue);
                                    break;
                                case CellType.Formula:
                                    rawCell.SetCellValue(targetWrap.RawCell.CellFormula);
                                    break;
                                case CellType.Blank:
                                    break;
                                case CellType.Boolean:
                                    rawCell.SetCellValue(targetWrap.RawCell.BooleanCellValue);
                                    break;
                                case CellType.Error:
                                    break;
                                default:
                                    rawCell.SetCellValue(targetWrap.Value);
                                    break;
                            }
                        }

                        tableModified = true;
                    }

                }
            }

            if (tableModified)
            {
                using (FileStream stream = new FileStream(rawFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream);
                }
            }
        }
    }

}
