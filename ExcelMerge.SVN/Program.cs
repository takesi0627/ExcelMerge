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
            var theirPath = args[1];
            var minePath = args[2];
            var basePath = args[3];

            // var theirPath = @"D:\their.xlsx";
            // var minePath = @"D:\mine.xlsx";
            // var basePath = @"D:\base.xlsx";

            var baseExcel = ExcelWorkbook.CreateSVN(basePath);
            var mineExcel = ExcelWorkbook.CreateSVN(minePath);
            var theirExcel = ExcelWorkbook.CreateSVN(theirPath);

            ExcelSheetDiffConfig config = new ExcelSheetDiffConfig();

            foreach (var name in baseExcel.SheetNames)
            {
                var mineSheetDiff = ExcelSheet.Diff(baseExcel.Sheets[name], mineExcel.Sheets[name], config);
                var theirSheetDiff = ExcelSheet.Diff(baseExcel.Sheets[name], theirExcel.Sheets[name], config);

                // 准备好一个空的 ExcelSheetDiff， finalDiff
                ExcelSheetDiff finalSheetDiff = new ExcelSheetDiff();
                

                // 准备三个index，一个记录分别记录三个 diff 读到哪了
                int finalIndex = 0;
                int mineIndex = 0;
                int theirIndex = 0;
                
                while (true)
                {
                    if (!mineSheetDiff.Rows.ContainsKey(mineIndex))
                    {
                        break;
                    }
                    // 首先从零开始，把 mineDiff 和 theirDiff 中左边为空的全部添加到 finalDiff
                    while (mineSheetDiff.Rows.ContainsKey(mineIndex) && mineSheetDiff.Rows[mineIndex].LeftEmpty())
                    {
                        finalSheetDiff.Rows.Add(finalIndex, mineSheetDiff.Rows[mineIndex]);
                        finalIndex++;
                        mineIndex++;
                    }

                    while (theirSheetDiff.Rows.ContainsKey(theirIndex) && theirSheetDiff.Rows[theirIndex].LeftEmpty())
                    {
                        finalSheetDiff.Rows.Add(finalIndex, theirSheetDiff.Rows[mineIndex]);
                        finalIndex++;
                        theirIndex++;
                    }

                    if (!mineSheetDiff.Rows.ContainsKey(mineIndex))
                    {
                        break;
                    }

                    // mine their 各进1，merge这个 rowDiff，放到 finalDiff（这里需要校验是不是冲突）
                    var mergedRowDiff = mineSheetDiff.Rows[mineIndex].getMergedRowDiff(theirSheetDiff.Rows[theirIndex]);

                    finalSheetDiff.Rows.Add(finalIndex, mergedRowDiff);
                    finalIndex++;
                    mineIndex++;
                    theirIndex++;
                }

                baseExcel.MergeSheetDiff(name, finalSheetDiff);
            }

            baseExcel.SaveAs(mergedPath);
        }
    }
}
