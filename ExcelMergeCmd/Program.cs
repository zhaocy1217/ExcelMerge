using ExcelMerge;
using System;
using System.Linq;
using CommandLineArgsHelpers;
using System.Collections.Generic;
using SimpleJSON;
using NPOI.SS.UserModel;
namespace ExcelMergeCmd
{
    internal class Program
    {
        static void JsonAddValue(JSONArray jsonCells, ICell cell)
        {
            if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
            {
                jsonCells.Add(cell.NumericCellValue);
            }
            else if (cell.CellType == NPOI.SS.UserModel.CellType.String)
            {
                jsonCells.Add(cell.StringCellValue);
            }
            else if (cell.CellType == NPOI.SS.UserModel.CellType.Boolean)
            {
                jsonCells.Add(cell.BooleanCellValue);
            }
            else if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
            {
                jsonCells.Add(cell.CellFormula);
            }
            else if (cell.CellType == NPOI.SS.UserModel.CellType.Error)
            {
                jsonCells.Add(cell.ErrorCellValue);
            }
            else
            {
                jsonCells.Add(cell.StringCellValue);
            }
        }

        static void Main(string[] args)
        {
            var arguments = new CommandLineArgsHelper(args);// -p="base.xlsx" -c="current.xlsx" -sheets="Sheet1,Sheet2"
            (string parentFile, string currentFile) = (arguments["p"], arguments["c"]);
            string sheetsToCompare = arguments["sheets"];
            if (string.IsNullOrEmpty(sheetsToCompare))
            {
                Console.WriteLine("No sheet to compare specified");
                return;
            }
            var sheetsToCompareList = sheetsToCompare.Split(',');
            var readConfig = new ExcelSheetReadConfig() { TrimLastBlankColumns = true, TrimLastBlankRows = true };
            var parentWorkbook = string.IsNullOrEmpty(parentFile) ? null : ExcelWorkbook.Create(parentFile, readConfig);
            var parentSheetNames = parentWorkbook == null ? new List<string>() : parentWorkbook.Sheets.Keys.ToList();
            var currentWorkbook = string.IsNullOrEmpty(currentFile) ? null : ExcelWorkbook.Create(currentFile, readConfig);
            if (currentWorkbook == null)
            {
                throw new Exception($"Current workbook is null {currentFile} open failed");
            }
            var currentWorkbookSheetNames = currentWorkbook == null ? new List<string>() : currentWorkbook.Sheets.Keys.ToList();
            bool IsNewSheet(string sheetName)
            {
                return parentWorkbook == null
                    || !parentSheetNames.Contains(sheetName);
            }
            ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();
            var allSheets = new JSONObject();
            JSONObject WriteJsonFromRow(ExcelSheetDiff sheetDiff, string sheetName, int rowIndex, ExcelCellStatus status)
            {
                var diffRow = sheetDiff.Rows[rowIndex];
                ExcelWorkbook workbook = null;
                if (status == ExcelCellStatus.Added || status == ExcelCellStatus.Modified)
                {
                    workbook = currentWorkbook;
                }
                else
                {
                    workbook = parentWorkbook;
                }
                var row_dict = new JSONObject();
                var arr = new JSONArray();
                row_dict[rowIndex.ToString()] = arr;
                foreach (var kv in diffRow.Cells)
                {
                    var diffCell = kv.Value;
                    var cell = workbook.GetCell(sheetName,
                        diffCell.DstCell.OriginalRowIndex, diffCell.DstCell.OriginalColumnIndex);
                    JsonAddValue(arr, cell);
                }
                return row_dict;
            }
            HashSet<int> insertedRowIndex = new HashSet<int>();
            foreach (var sheetName in sheetsToCompareList)
            {
                if (!currentWorkbookSheetNames.Contains(sheetName))
                {
                    Console.WriteLine($"Sheet {sheetName} not found in current workbook");
                    continue;
                }
                var allChgedRows = new JSONArray();
                allSheets[sheetName] = allChgedRows;
                if (IsNewSheet(sheetName))
                {
                    var insertedRows = new JSONArray();
                    allChgedRows.Add(insertedRows);
                    var sheet = currentWorkbook.Sheets[sheetName];
                    foreach (var row_kv in sheet.Rows)
                    {
                        if (insertedRowIndex.Contains(row_kv.Key))
                            continue;
                        insertedRowIndex.Add(row_kv.Key);
                        var jsonCells = new JSONArray();
                        var row_dict = new JSONObject();
                        row_dict[row_kv.Key.ToString()] = jsonCells;
                        insertedRows.Add(row_dict);
                        foreach (var cell_kv in row_kv.Value.Cells)
                        {
                            var cell = currentWorkbook.GetCell(sheetName,
                                cell_kv.OriginalRowIndex, cell_kv.OriginalColumnIndex);
                            JsonAddValue(jsonCells, cell);
                        }
                    }
                }
                else
                {
                    var mineDiff = ExcelSheet.Diff(ExcelWorkbook.GetSheet(parentFile, sheetName, readConfig)
                                     , ExcelWorkbook.GetSheet(currentFile, sheetName, readConfig), diffConfig);
                    var mineSummary = mineDiff.CreateSummary();
                    if (mineSummary.InsertedRowIndexes.Count > 0)
                    {
                        foreach (var index in mineSummary.InsertedRowIndexes)
                        {
                            if (insertedRowIndex.Contains(index))
                                continue;
                            insertedRowIndex.Add(index);
                            allChgedRows.Add(WriteJsonFromRow(mineDiff, sheetName, index, ExcelCellStatus.Added));
                        }
                    }
                    if (mineSummary.ModifiedRowIndexes.Count > 0)
                    {
                        foreach (var index in mineSummary.ModifiedRowIndexes)
                        {
                            if (insertedRowIndex.Contains(index))
                                continue;
                            insertedRowIndex.Add(index);
                            allChgedRows.Add(WriteJsonFromRow(mineDiff, sheetName, index, ExcelCellStatus.Modified));
                        }
                    }
                }
                allChgedRows.Sort(new JsonArrayRowCompare());
            }
            var jsonStr = allSheets.ToString();
            Console.WriteLine(jsonStr);
        }
    }
    class JsonArrayRowCompare : IComparer<JSONNode>
    {
        public int Compare(JSONNode x, JSONNode y)
        {
            var x_keys = x.GetEnumerator();
            var y_keys = y.GetEnumerator();
            x_keys.MoveNext();
            y_keys.MoveNext();
            return int.Parse(x_keys.Current.Key) - int.Parse(y_keys.Current.Key);
        }
    }
}
