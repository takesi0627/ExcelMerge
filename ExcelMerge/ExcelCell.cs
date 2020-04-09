using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }
        public ICellStyle OriginalCellStyle { get; private set; }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex, ICellStyle style = null)
        {
            Value = value;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
            OriginalCellStyle = style;
        }
    }
}
