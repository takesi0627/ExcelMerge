using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NetDiff;
using SKCore.Collection;

namespace ExcelMerge
{
    public class ExcelSheet
    {
        public SortedDictionary<int, ExcelRow> Rows { get; private set; }

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

        public static ExcelSheet Create(ISheet srcSheet, ExcelSheetReadConfig config)
        {
            var rows = ExcelReader.Read(srcSheet);

            return CreateSheet(rows, config);
        }

        public static ExcelSheet CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var rows = CsvReader.Read(path);

            return CreateSheet(rows, config);
        }

        public static ExcelSheet CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var rows = TsvReader.Read(path);

            return CreateSheet(rows, config);
        }

        private static ExcelSheet CreateSheet(IEnumerable<ExcelRow> rows, ExcelSheetReadConfig config)
        {
            return new ExcelSheet(rows, config);
        }

        public void SetCell(int row, int col, string value)
        {
            Rows[row].Cells[col].Value = value;
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

        public static ExcelSheetDiff Diff(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config)
        {
            var srcColumns = src.CreateColumns();
            var dstColumns = dst.CreateColumns();
            var columnStatusMap = CreateColumnStatusMap(srcColumns, dstColumns, config);

            var option = new DiffOption<ExcelRow>();
            option.EqualityComparer =
                new RowComparer(new HashSet<int>(columnStatusMap.Where(i => i.Value != ExcelColumnStatus.None).Select(i => i.Key)));

            // 这里实际上计算的是有没有插入新列
            foreach (var row in src.Rows.Values)
            {
                var shifted = new List<ExcelCell>();
                var index = 0;
                var queue = new Queue<ExcelCell>(row.Cells);
                while (queue.Any())
                {
                    if (columnStatusMap[index] == ExcelColumnStatus.Inserted)
                        shifted.Add(new ExcelCell(string.Empty, 0, 0));
                    else
                        shifted.Add(queue.Dequeue());

                    index++;
                }

                row.UpdateCells(shifted);
            }

            foreach (var row in dst.Rows.Values)
            {
                var shifted = new List<ExcelCell>();
                var index = 0;
                var queue = new Queue<ExcelCell>(row.Cells);
                while (queue.Any())
                {
                    if (columnStatusMap[index] == ExcelColumnStatus.Deleted)
                        shifted.Add(new ExcelCell(string.Empty, 0, 0));
                    else
                        shifted.Add(queue.Dequeue());

                    index++;
                }

                row.UpdateCells(shifted);
            }

            var r = DiffUtil.Diff(src.Rows.Values, dst.Rows.Values, option);
            r = DiffUtil.Order(r, DiffOrderType.LazyDeleteFirst);
            var resultArray = DiffUtil.OptimizeCaseDeletedFirst(r).ToArray();
            if (resultArray.Length > 10000)
            {
                var count = 0;
                var indices = Enumerable.Range(0, 100).ToList();
                foreach (var result in resultArray)
                {
                    if (result.Status != DiffStatus.Equal)
                        indices.AddRange(Enumerable.Range(Math.Max(0, count - 100), 200));

                    count++;
                }
                indices = indices.Distinct().ToList();
                resultArray = indices.Where(i => i < resultArray.Length).Select(i => resultArray[i]).ToArray();
            }

            var sheetDiff = new ExcelSheetDiff();
            DiffCells(resultArray, sheetDiff, columnStatusMap);

            return sheetDiff;
        }

        private static Dictionary<int, ExcelColumnStatus> CreateColumnStatusMap(
            IEnumerable<ExcelColumn> srcColumns, IEnumerable<ExcelColumn> dstColumns, ExcelSheetDiffConfig config)
        {
            var option = new DiffOption<ExcelColumn>();

            if (config.SrcHeaderIndex >= 0)
            {
                option.EqualityComparer = new HeaderComparer();
                foreach (var sc in srcColumns)
                    sc.HeaderIndex = config.SrcHeaderIndex;
            }

            if (config.DstHeaderIndex >= 0)
            {
                foreach (var dc in dstColumns)
                    dc.HeaderIndex = config.DstHeaderIndex;
            }

            var results = DiffUtil.Diff(srcColumns, dstColumns, option);
            results = DiffUtil.Order(results, DiffOrderType.LazyDeleteFirst);
            results = DiffUtil.OptimizeCaseDeletedFirst(results);
            var ret = new Dictionary<int, ExcelColumnStatus>();
            var columnIndex = 0;
            foreach (var result in results)
            {
                var status = ExcelColumnStatus.None;
                if (result.Status == DiffStatus.Deleted)
                    status = ExcelColumnStatus.Deleted;
                else if (result.Status == DiffStatus.Inserted)
                    status = ExcelColumnStatus.Inserted;

                ret.Add(columnIndex, status);
                columnIndex++;
            }

            return ret;
        }

        private IEnumerable<ExcelColumn> CreateColumns()
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

        private static void DiffCells(
            IEnumerable<DiffResult<ExcelRow>> results, ExcelSheetDiff sheetDiff, Dictionary<int, ExcelColumnStatus> columnStatusMap)
        {
            foreach (var result in results)
            {
                switch (result.Status)
                {
                    case DiffStatus.Equal:
                        DiffCellsCaseEqual(result, sheetDiff, columnStatusMap);
                        break;
                    case DiffStatus.Modified:
                        DiffCellsCaseEqual(result, sheetDiff, columnStatusMap);
                        break;
                    case DiffStatus.Deleted:
                        DiffCellsCaseDeleted(result, sheetDiff, columnStatusMap);
                        break;
                    case DiffStatus.Inserted:
                        DiffCellsCaseInserted(result, sheetDiff, columnStatusMap);
                        break;
                }
            }
        }

        private static IEnumerable<Tuple<ExcelCell, ExcelCell>> EqualizeColumnCount(
            IEnumerable<ExcelCell> srcCells, IEnumerable<ExcelCell> dstCells, Dictionary<int, ExcelColumnStatus> columnStausMap)
        {
            var srcQueue = new Queue<ExcelCell>(srcCells);
            var dstQueue = new Queue<ExcelCell>(dstCells);
            foreach (var status in columnStausMap)
            {
                ExcelCell src = null;
                ExcelCell dst = null;

                if (srcQueue.Any()) src = srcQueue.Dequeue();
                if (dstQueue.Any()) dst = dstQueue.Dequeue();

                yield return Tuple.Create(src, dst);
            }
        }

        private static void DiffCellsCaseEqual(
            DiffResult<ExcelRow> result, ExcelSheetDiff sheetDiff, Dictionary<int, ExcelColumnStatus> columnStatusMap)
        {
            var row = sheetDiff.CreateRow();

            var equalizedCells = EqualizeColumnCount(result.Obj1.Cells, result.Obj2.Cells, columnStatusMap);
            var columnIndex = 0;
            foreach (var pair in equalizedCells)
            {
                var srcCell = pair.Item1;
                var dstCell = pair.Item2;

                if (srcCell != null && dstCell != null)
                {
                    var status = srcCell.Value.Equals(dstCell.Value) ? ExcelCellStatus.None : ExcelCellStatus.Modified;
                    if (columnStatusMap[columnIndex] == ExcelColumnStatus.Deleted)
                        status = ExcelCellStatus.Removed;
                    else if (columnStatusMap[columnIndex] == ExcelColumnStatus.Inserted)
                        status = ExcelCellStatus.Added;

                    row.CreateCell(srcCell, dstCell, columnIndex, status);
                }
                else if (srcCell != null && dstCell == null)
                {
                    dstCell = new ExcelCell(string.Empty, srcCell.OriginalColumnIndex, srcCell.OriginalColumnIndex);
                    row.CreateCell(srcCell, dstCell, columnIndex, ExcelCellStatus.Removed);
                }
                else if (srcCell == null && dstCell != null)
                {
                    srcCell = new ExcelCell(string.Empty, dstCell.OriginalColumnIndex, dstCell.OriginalColumnIndex);
                    row.CreateCell(srcCell, dstCell, columnIndex, ExcelCellStatus.Added);
                }
                else
                {
                    srcCell = new ExcelCell(string.Empty, 0, 0);
                    dstCell = new ExcelCell(string.Empty, 0, 0);
                    row.CreateCell(srcCell, dstCell, columnIndex, ExcelCellStatus.None);
                }

                columnIndex++;
            }
        }

        private static void DiffCellsCaseDeleted(
            DiffResult<ExcelRow> result, ExcelSheetDiff sheetDiff, Dictionary<int, ExcelColumnStatus> columnStatusMap)
        {
            var row = sheetDiff.CreateRow();

            var columnIndex = 0;
            foreach (var cell1 in result.Obj1.Cells)
            {
                var cell2 = new ExcelCell(string.Empty, cell1.OriginalColumnIndex, cell1.OriginalRowIndex);
                row.CreateCell(cell1, cell2, columnIndex, ExcelCellStatus.Removed);

                columnIndex++;
            }
        }

        private static void DiffCellsCaseInserted(
            DiffResult<ExcelRow> result, ExcelSheetDiff sheetDiff, Dictionary<int, ExcelColumnStatus> columnStatusMap)
        {
            var row = sheetDiff.CreateRow();

            var columnIndex = 0;
            foreach (var cell2 in result.Obj2.Cells)
            {
                var cell1 = new ExcelCell(string.Empty, cell2.OriginalColumnIndex, cell2.OriginalRowIndex);
                row.CreateCell(cell1, cell2, columnIndex, ExcelCellStatus.Added);

                columnIndex++;
            }
        }
    }
}
