using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI;
namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }
        public IWorkbook RawWorkbook { get; private set; }
        public string ExcelPath { get; private set; }
        public ExcelWorkbook(string path)
        {
            Sheets = new Dictionary<string, ExcelSheet>();
            ExcelPath = path;
        }
        public void InsertColumn(string sheetName, int index, ExcelColumn column, List<int> newInsertedRows, Action<ICell> onCellCreatAct)
        {
            Sheets[sheetName].InsertColumn(index, newInsertedRows, column);
            var sheet = RawWorkbook.GetSheet(sheetName);
            ExcelUtility.InsertColumn(sheet, index, newInsertedRows, onCellCreatAct);
        }
        public void InsertRow(string sheetName, int index, ExcelRow row, Action<ICell, int> onCellCreatAct)
        {
            var sheet = RawWorkbook.GetSheet(sheetName);
            int lastRowNum = sheet.LastRowNum;
            if (index <= lastRowNum)
            {
                sheet.ShiftRows(index, lastRowNum, 1);
            }
            IRow newRow = sheet.CreateRow(index);
            for (int i = 0; i < row.Cells.Count; i++)
            {
                var newCell = newRow.CreateCell(i);
                onCellCreatAct(newCell, i);
            }
        }
        public void Save()
        {
            var workbook = RawWorkbook;
            try
            {
                using (FileStream outputStream = new FileStream(ExcelPath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(outputStream);
                }
                workbook.Close();
                System.Console.WriteLine($"Successfully wrote changes to {ExcelPath}");
            }
            catch (IOException ex)
            {
                System.Console.WriteLine($"Error writing to file: {ex.Message}. Make sure the file is not open or in use by another process.");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
        public ICell GetCell(string sheetName, int rowIndex, int columnIdx)
        {
            var sheet = RawWorkbook.GetSheet(sheetName);
            var row = sheet.GetRow(rowIndex);
            var cell = row.GetCell(columnIdx);
            return cell;
        }
        public static IWorkbook WorkbookFactoryNewCreate(string path)
        {
            Stream inputStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            if (inputStream.Length == 0)
                throw new Exception("File is empty");
            inputStream = new PushbackStream(inputStream);
            if (POIFSFileSystem.HasPOIFSHeader(inputStream))
            {
                return new NPOI.HSSF.UserModel.HSSFWorkbook(inputStream);
            }
            inputStream.Position = 0;
            if (ExcelUtility.HasOOXMLHeader(inputStream))
            {
                var pkg = PackageHelper.Open(inputStream, false);
                return new XSSFWorkbook(pkg);
            }
            throw new InvalidFormatException("Your stream was neither an OLE2 stream, nor an OOXML stream.");
        }
        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);
            var srcWb = WorkbookFactoryNewCreate(path);
            var wb = new ExcelWorkbook(path);
            wb.RawWorkbook = srcWb;
            for (int i = 0; i < srcWb.NumberOfSheets; i++)
            {
                var srcSheet = srcWb.GetSheetAt(i);
                wb.Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
            }

            return wb;
        }

        public static ExcelSheet GetSheet(string path, string sheetName, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                throw new Exception("CSV file is not supported");

            if (Path.GetExtension(path) == ".tsv")
                throw new Exception("TSV file is not supported");

            var srcWb = WorkbookFactoryNewCreate(path);
            var srcSheet = srcWb.GetSheet(sheetName);
            var sheet = ExcelSheet.Create(srcSheet, config);
            return sheet;
        }
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
                var wb = WorkbookFactoryNewCreate(path);
                for (int i = 0; i < wb.NumberOfSheets; i++)
                    yield return wb.GetSheetAt(i).SheetName;
            }
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook(path);
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromCsv(path, config));

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook(path);
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromTsv(path, config));

            return wb;
        }
    }
}
