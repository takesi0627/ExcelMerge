using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelSheetDiff
    {
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }

        public ExcelSheetDiff()
        {
            Rows = new SortedDictionary<int, ExcelRowDiff>();
        }

        public ExcelRowDiff CreateRow()
        {
            var row = new ExcelRowDiff(Rows.Any() ? Rows.Keys.Last() + 1 : 0);
            Rows.Add(row.Index, row);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            foreach (var row in Rows)
            {
                if (row.Value.IsAdded())
                    addedRowCount++;
                else if (row.Value.IsRemoved())
                    removedRowCount++;

                if (row.Value.IsModified())
                    modifiedRowCount++;

                modifiedCellCount += row.Value.ModifiedCellCount;
            }

            return new ExcelSheetDiffSummary
            {
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
            };
        }

        public ExcelCellDiff GetCell(int row, int col)
        {
            if (!Rows.ContainsKey(row) || !Rows[row].Cells.ContainsKey(col))
                return null;

            return Rows[row].Cells[col];
        }

        public void Merge(IEnumerable<int> rows, ExcelCellMergeStatus mergeStatus)
        {
            foreach (var row in rows)
            {
                Merge(row, mergeStatus);
            }
        }

        public void Merge(int row, ExcelCellMergeStatus mergeStatus)
        {
            if (!Rows.ContainsKey(row))
                return;

            Rows[row].Merge(mergeStatus);
        }
    }
}
