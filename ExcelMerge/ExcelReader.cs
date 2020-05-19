using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

// using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(IXLWorksheet sheet)
        {
            var wb = new XLWorkbook();

            int totalRowCount = sheet.LastRowUsed().RowNumber();

            for (int nowRowIndex = 1; nowRowIndex < totalRowCount+1; nowRowIndex++)
            {
                var nowRow = sheet.Row(nowRowIndex);
                if (nowRow == null)
                    continue;

                var cells = new List<ExcelCell>();

                var totalCellCount = nowRow.LastCellUsed().Address.ColumnNumber;

                for (int nowCellIndex = 1; nowCellIndex < totalCellCount+1; nowCellIndex++)
                {
                    var cell = nowRow.Cell(nowCellIndex);
                    if (cell != null)
                    {
                        if (cell.HasFormula)
                        {
                            var formulaString = cell.FormulaA1;
                            cells.Add(new ExcelCell(formulaString, nowCellIndex, nowRowIndex, cell));
                        }
                        else
                        {
                            cells.Add(new ExcelCell(cell.GetString(), nowCellIndex, nowRowIndex, cell));
                        }
                    }
                    else
                    {
                        cells.Add(new ExcelCell("", nowCellIndex, nowRowIndex, cell));
                    }
                }

                yield return new ExcelRow(nowRowIndex, cells);
            }

           
        }
    }
}
