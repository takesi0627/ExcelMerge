using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }
        public ICell RawCell { get; }


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
    }
}
