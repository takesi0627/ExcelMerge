// using NPOI.SS.UserModel;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }
        public IXLCell RawCell { get; }


        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex, IXLCell rawCell = null)
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

        public bool ValueEqual(ExcelCell cell)
        {
            return Value == cell.Value;
        }
    }
}
