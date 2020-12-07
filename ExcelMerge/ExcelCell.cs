using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }
        public ICell RawCell { get => _rawCell; set => _rawCell = value; }
        private ICell _rawCell;

        public bool IsDirty { get; private set; } = false;

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex, ICell rawCell = null)
        {
            Value = value;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
            RawCell = rawCell;
        }

        public ExcelCell Clone()
        {
            return new ExcelCell(Value, OriginalColumnIndex, OriginalRowIndex, RawCell);
        }

        public void OverwriteValueBy(ExcelCell src)
        {
            Value = src.Value;
            ExcelUtility.CloneRawCell(src.RawCell, ref _rawCell);

            IsDirty = true;
        }
    }
}
