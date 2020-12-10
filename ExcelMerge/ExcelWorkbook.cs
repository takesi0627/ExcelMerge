using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace ExcelMerge
{
    public abstract class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; } = new Dictionary<string, ExcelSheet>();

        public List<string> SheetNames { get => Sheets.Keys.ToList(); }

        protected string rawFilePath;

        public ExcelWorkbook(string path)
        {
            rawFilePath = path;
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            var ext = Path.GetExtension(path);
            if (ext == ".csv" || ext == ".tsv")
            {
                return new SVWorkbook(path, config);
            }
            else if (ext == ".xlsx" || ext == ".xlsm")
            {
                return new XLSWorkbook(path, config);
            }
            else
            {
                Debug.Assert(false, $"Invalid file type: {ext}");
                return null;
            }
        }

        public abstract void Save();

        public static IEnumerable<string> GetSheetNames(string path)
        {
            if (Path.GetExtension(path) == ".csv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else if (Path.GetExtension(path) == ".tsv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else
            {
                var wb = WorkbookFactory.Create(path);
                for (int i = 0; i < wb.NumberOfSheets; i++)
                    yield return wb.GetSheetAt(i).SheetName;
            }
        }
    }

    internal class SVWorkbook : ExcelWorkbook
    {
        internal SVWorkbook(string path, ExcelSheetReadConfig config) : base (path)
        {
            Sheets.Add(Path.GetFileName(path), new SVExcelSheet(path, config));
        }

        public override void Save()
        {
            var filename = Path.GetFileName(rawFilePath);

            var sheet = Sheets[filename] as SVExcelSheet;
            // skip if this sheet is not edited
            if (!sheet.IsDirty)
            {
                return;
            }

            var ext = Path.GetExtension(rawFilePath);
            var sep = ext == ".csv" ? "," : "\t";
            var nl = System.Environment.NewLine;
            var maxRow = sheet.Rows.Keys.Max();

            string output = "";
            foreach (var row in sheet.Rows.Values)
            {
                output += string.Join(sep, row.Cells.Select(c => c.Value)) + nl;
            }

            using (StreamWriter stream = new StreamWriter(rawFilePath, false, Encoding.UTF8))
            {
                stream.Write(output);
            }

            sheet.IsDirty = false;
        }
    }

    internal class XLSWorkbook : ExcelWorkbook
    {
        private XSSFWorkbook rawWorkbook;

        internal XLSWorkbook(string path, ExcelSheetReadConfig config) : base (path)
        {
            string tmpFile = Path.GetTempFileName();
            File.Copy(path, tmpFile, true);

            rawWorkbook = new XSSFWorkbook(tmpFile);

            for (int i = 0; i < rawWorkbook.NumberOfSheets; i++)
            {
                var srcSheet = rawWorkbook.GetSheetAt(i);
                Sheets.Add(srcSheet.SheetName, new XLSExcelSheet(srcSheet, config));
            }
        }

        public override void Save()
        {
            if (Sheets.Any(s => s.Value.IsDirty))
            {
                using (FileStream stream = new FileStream(rawFilePath, FileMode.Create, FileAccess.Write))
                {
                    rawWorkbook.Write(stream);
                }

                foreach (var sheet in Sheets.Values)
                {
                    sheet.IsDirty = false;
                }
            }
        }
    }
}
