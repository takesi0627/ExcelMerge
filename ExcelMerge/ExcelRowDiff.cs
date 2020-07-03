using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelMerge
{
    public class ExcelRowDiff
    {
        public int Index { get; private set; }
        public SortedDictionary<int, ExcelCellDiff> Cells { get; private set; }

        public ExcelRowDiff(int index)
        {
            Index = index;
            Cells = new SortedDictionary<int, ExcelCellDiff>();
        }

        public ExcelCellDiff CreateCell(ExcelCell src, ExcelCell dst, int columnIndex, ExcelCellStatus status)
        {
            var cell = new ExcelCellDiff(columnIndex, Index, src, dst, status);
            Cells.Add(cell.ColumnIndex, cell);

            return cell;
        }

        public bool IsModified()
        {
            return Cells.Any(c => c.Value.Status != ExcelCellStatus.None);
        }

        public bool IsAdded()
        {
            return Cells.All(c => c.Value.Status == ExcelCellStatus.Added);
        }

        public bool IsRemoved()
        {
            return Cells.All(c => c.Value.Status == ExcelCellStatus.Removed);
        }

        public bool NeedMerge()
        {
            return Cells.Any(c => c.Value.MergeStatus != ExcelCellMergeStatus.None);
        }

        public bool LeftEmpty()
        {
            return Cells.All(c => c.Value.SrcCell.Value == string.Empty);
        }

        public bool RightEmpty()
        {
            return Cells.All(c => c.Value.DstCell.Value == string.Empty);
        }

        public int ModifiedCellCount
        {
            get { return Cells.Count(c => c.Value.Status != ExcelCellStatus.None); }
        }

        public bool LeftEqual(ExcelRowDiff otherDiff)
        {
            foreach (var cellDiff in Cells)
            {
                if (!otherDiff.Cells.ContainsKey(cellDiff.Key))
                {
                    return false;
                }

                if (!cellDiff.Value.SrcCell.ValueEqual(otherDiff.Cells[cellDiff.Key].SrcCell))
                {
                    return false;
                }
            }

            return true;
        }

        public ExcelRowDiff getMergedRowDiff(ExcelRowDiff otherDiff)
        {
            Debug.Assert(LeftEqual(otherDiff));
            var mergedRowDiff = new ExcelRowDiff(-1);

            int cellIndex = 0;

            foreach (var cellDiff in Cells)
            {
                var cellKey = cellDiff.Key;

                if (cellDiff.Value.Status == ExcelCellStatus.None)
                {
                    mergedRowDiff.Cells.Add(cellIndex, otherDiff.Cells[cellKey]);
                }
                else
                {
                    Debug.Assert(otherDiff.Cells[cellKey].Status == ExcelCellStatus.None);
                    mergedRowDiff.Cells.Add(cellIndex, cellDiff.Value);
                }

                cellIndex++;
            }

            return mergedRowDiff;
        }

        // TODO: Add row status field and implemnt UpdateStaus method.
    }
}
