using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NetDiff;

namespace ExcelMerge
{
    public abstract class ExcelSheet
    {
        public SortedDictionary<int, ExcelRow> Rows { get; private set; }

        public bool IsDirty = false;

        public ExcelSheet(IEnumerable<ExcelRow> rows, ExcelSheetReadConfig config)
        {
            Rows = new SortedDictionary<int, ExcelRow>();

            foreach (var row in rows)
            {
                Rows.Add(row.Index, row);
            }

            if (config.TrimFirstBlankRows)
                TrimFirstBlankRows();

            if (config.TrimFirstBlankColumns)
                TrimFirstBlankColumns();

            if (config.TrimLastBlankRows)
                TrimLastBlankRows();

            if (config.TrimLastBlankColumns)
                TrimLastBlankColumns();
        }

        public abstract string GetCellValue(int row, int col);

        public virtual void SetCell(int row, int col, string value)
        {
            if (!Rows.ContainsKey(row))
            {
                var cnt = Rows.Select(r => r.Value.Cells.Count).Max();
                var list = Enumerable.Range(0, cnt).Select(x => new ExcelCell("", x, row));

                Rows.Add(row, new ExcelRow(row, list));
            }
            else if (Rows[row].Cells.Count < col + 1)
            {
                for (var c = Rows[row].Cells.Count; c < col + 1; ++c)
                {
                    Rows[row].Cells.Add(new ExcelCell("", c, row));
                }
            }

            Rows[row].Cells[col].Value = value;
            IsDirty = true;
        }

        public void TrimFirstBlankRows()
        {
            var rows = new SortedDictionary<int, ExcelRow>();
            var index = 0;
            foreach (var row in Rows.SkipWhile(r => r.Value.IsBlank()))
            {
                rows.Add(index, new ExcelRow(index, row.Value.Cells));
                index++;
            }

            Rows = rows;
        }

        public void TrimFirstBlankColumns()
        {
            var columns = CreateColumns();
            var indices = columns.Select((v, i) => new { v, i }).TakeWhile(c => c.v.IsBlank()).Select(c => c.i);

            foreach (var i in indices)
                RemoveColumn(i);
        }

        public void TrimLastBlankRows()
        {
            var rows = new SortedDictionary<int, ExcelRow>();
            var index = 0;
            foreach (var row in Rows.Reverse().SkipWhile(r => r.Value.IsBlank()).Reverse())
            {
                rows.Add(index, new ExcelRow(index, row.Value.Cells));
                index++;
            }

            Rows = rows;
        }

        public void TrimLastBlankColumns()
        {
            var columns = CreateColumns();
            var indices = columns.Select((v, i) => new { v, i }).Reverse().TakeWhile(c => c.v.IsBlank()).Select(c => c.i);

            foreach (var i in indices)
                RemoveColumn(i);
        }

        public void RemoveColumn(int column)
        {
            foreach (var row in Rows)
            {
                if (row.Value.Cells.Count > column)
                    row.Value.Cells.RemoveAt(column);
            }
        }

        internal IEnumerable<ExcelColumn> CreateColumns()
        {
            if (!Rows.Any())
                return Enumerable.Empty<ExcelColumn>();

            var columnCount = Rows.Max(r => r.Value.Cells.Count);
            var columns = new ExcelColumn[columnCount];
            foreach (var row in Rows)
            {
                var columnIndex = 0;
                foreach (var cell in row.Value.Cells)
                {
                    if (columns[columnIndex] == null)
                        columns[columnIndex] = new ExcelColumn();

                    columns[columnIndex].Cells.Add(cell);
                    columnIndex++;
                }
            }

            return columns.AsEnumerable();
        }
    }

    /// <summary>
    /// Sheet data for .xlsl/.xlsm file
    /// </summary>
    internal class XLSExcelSheet : ExcelSheet
    {
        private ISheet rawSheet;

        public XLSExcelSheet(ISheet originalSheet, ExcelSheetReadConfig config) : base(ExcelReader.Read(originalSheet), config)
        {
            rawSheet = originalSheet;
        }

        public override void SetCell(int row, int col, string value)
        {
            base.SetCell(row, col, value);

            IRow rawRow = rawSheet.GetRow(row);
            if (rawRow == null)
            {
                rawRow = rawSheet.CreateRow(row);
            }

            ICell rawCell = rawRow.GetCell(col);
            if (rawCell == null)
            {
                rawCell = rawRow.CreateCell(col);
            }

            rawCell.SetCellValue(value);
        }

        public override string GetCellValue(int row, int col)
        {
            return ExcelUtility.GetCellStringValue(rawSheet.GetRow(row).GetCell(col));
        }
    }

    /// <summary>
    /// Sheet data for .csv/.tsv file
    /// </summary>
    internal class SVExcelSheet : ExcelSheet
    {
        public SVExcelSheet(string path, ExcelSheetReadConfig config) : base(ExcelReader.Read(path), config)
        {
        }

        public override string GetCellValue(int row, int col)
        {
            return Rows[row].Cells[col].Value;
        }
    }
}
