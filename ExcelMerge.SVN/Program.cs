using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge.SVN
{
    class Program
    {
        static void Main(string[] args)
        {
            var mergedPath = args[0];
            var theirPath = @"D:\their.xlsx";
            var minePath = @"D:\mine.xlsx"; ;
            var basePath = @"D:\base.xlsx"; ;

            var baseExcel = ExcelWorkbook.CreateSVN(basePath);
            var mineExcel = ExcelWorkbook.CreateSVN(minePath);
            var theirExcel = ExcelWorkbook.CreateSVN(theirPath);

            ExcelSheetDiffConfig config = new ExcelSheetDiffConfig();

            foreach (var name in baseExcel.SheetNames)
            {
                var mineDiff = ExcelSheet.Diff(baseExcel.Sheets[name], mineExcel.Sheets[name], config);
                var theirDiff = ExcelSheet.Diff(baseExcel.Sheets[name], theirExcel.Sheets[name], config);
            }
        }
    }
}
