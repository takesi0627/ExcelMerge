using System;
using ExcelMerge;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var targetPath = args[0];
            var theirPath = @"D:\their.xlsx";
            var minePath = @"D:\mine.xlsx"; ;
            var basePath = @"D:\base.xlsx"; ;

            var baseExcel = ExcelWorkbook.CreateSVN(basePath);
            var mineExcel = ExcelWorkbook.CreateSVN(minePath);
            var theirExcel = ExcelWorkbook.CreateSVN(theirPath);
        }
    }
}
