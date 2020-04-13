namespace ExcelMerge
{
    public class ExcelCellDiff
    {
        public int ColumnIndex { get; }
        public int RowIndex { get; }
        public ExcelCell SrcCell { get; }
        public ExcelCell DstCell { get; }
        public ExcelCellStatus Status { get; set; }
        public ExcelCellMergeStatus MergeStatus { get; set; }


        public ExcelCellDiff(int columnIndex, int rowIndex, ExcelCell src, ExcelCell dst, ExcelCellStatus status, ExcelCellMergeStatus mergeStatus = ExcelCellMergeStatus.None)
        {
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            SrcCell = src;
            DstCell = dst;
            Status = status;
            MergeStatus = mergeStatus;
        }

        public ExcelCellDiff Clone()
        {
            return new ExcelCellDiff(ColumnIndex, RowIndex, SrcCell.Clone(), DstCell.Clone(), Status, MergeStatus);
        }

        public override string ToString()
        {
            return $"Src: {SrcCell.Value} Dst: {DstCell.Value}: Status: {Status}";
        }
    }
}
