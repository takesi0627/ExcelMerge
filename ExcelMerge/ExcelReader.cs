using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var actualRowIndex = 0;
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                    continue;

                var cells = row.Cells.Select(cell =>
                {
                    var stringValue = ExcelUtility.GetCellStringValue(cell);
                    return new ExcelCell(stringValue, cell.ColumnIndex, cell.RowIndex, cell);
                });

                yield return new ExcelRow(actualRowIndex++, cells);
            }
        }

        internal static IEnumerable<ExcelRow> Read(string path)
        {
            var ext = Path.GetExtension(path);
            if (ext == ".csv")
            {
                return CsvReader.Read(path);
            }
            else if (ext == ".tsv")
            {
                return CsvReader.Read(path);
            }
            else
            {
                return new List<ExcelRow>();
            }
        }
    }
}
