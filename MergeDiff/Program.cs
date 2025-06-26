using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelMerge;
using System.Text.RegularExpressions;
using NetDiff;
using NPOI.HSSF.Record.PivotTable;
using System.Diagnostics;
using NPOI.SS.Formula.Functions;
using System.IO;

using NPOI.SS.UserModel;

namespace MergeDiff
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Test");
            (string parentFile, string othersFile, string mineFile) = (args[0], args[1], args[2]);
            Console.WriteLine($"Loading File {parentFile}");
            var readConfig = new ExcelSheetReadConfig() { TrimLastBlankColumns = true, TrimLastBlankRows = true };
            var parentWorkbook = ExcelWorkbook.Create(parentFile, readConfig);
            var parentSheetNames = parentWorkbook.Sheets.Keys.ToList();
            Console.WriteLine($"Loading File {othersFile}");
            var othersWorkbook = ExcelWorkbook.Create(othersFile, readConfig);
            var othersSheetNames = othersWorkbook.Sheets.Keys.ToList();
            Console.WriteLine($"Loading File {mineFile}");
            var mineWorkbook = ExcelWorkbook.Create(mineFile, readConfig);
            var mineSheetNames = mineWorkbook.Sheets.Keys.ToList();
            if (MergeDiffUtility.GetHashCode(parentSheetNames) != MergeDiffUtility.GetHashCode(othersSheetNames)
                || MergeDiffUtility.GetHashCode(parentSheetNames) != MergeDiffUtility.GetHashCode(mineSheetNames))
            {
                Console.WriteLine("Sheet names are different");
                return;
            }
            ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();
            bool NeedSimpleDiffCalc(ExcelSheet sheet)
            {
                var firstRowCellCount = sheet.Rows[diffConfig.SrcHeaderIndex].Cells.Count;
                var invalidCount = sheet.Rows.Count(row => row.Value.Cells.Count > firstRowCellCount);
                if (invalidCount > 0)
                {
                    return true;
                }
                return false;
            }
            bool IsValidSheet(ExcelSheet sheet)
            {
                var idCells = sheet.GetColumn(0).Cells.Count();
                if (idCells != sheet.Rows.Count)
                {
                    return false;
                }
                //for (int i = 0; i < sheet.MaxColumnCount; i++)
                //{
                //    var currentColumnRowCount = sheet.GetColumn(i).Cells.Count();
                //    if (currentColumnRowCount > idCells)
                //    {
                //        return false;
                //    }
                //}
                return true;
            }
            bool IsSimpleDiff(ExcelSheet sheet1, ExcelSheet sheet2)
            {
                foreach (var row in sheet1.Rows)
                {
                    if (row.Value.Cells.Count != sheet2.Rows[row.Key].Cells.Count)
                    {
                        return true;
                    }
                    for (int i = 0; i < row.Value.Cells.Count; i++)
                    {
                        if (row.Value.Cells[i].Value != sheet2.Rows[row.Key].Cells[i].Value)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }

            foreach (var sheetName in parentSheetNames)
            {
                if (MergeDiffUtility.ContainsChinese(sheetName))
                {
                    continue;
                }
                var parentSheet = parentWorkbook.Sheets[sheetName];
                if (parentSheet.Rows.Count == 0)
                {
                    continue;
                }
                var othersSheet = othersWorkbook.Sheets[sheetName];
                var mineSheet = mineWorkbook.Sheets[sheetName];
                if (!IsValidSheet(parentSheet) || !IsValidSheet(othersSheet) || !IsValidSheet(mineSheet))
                {
                    Console.WriteLine($"Sheet: {sheetName} is not valid");
                    continue;
                }
                //bool simpleDiff = NeedSimpleDiffCalc(parentSheet) || NeedSimpleDiffCalc(othersSheet) || NeedSimpleDiffCalc(mineSheet);
                //if (simpleDiff)
                //{
                //    if (IsSimpleDiff(parentSheet, othersSheet) || IsSimpleDiff(parentSheet, mineSheet))
                //    {
                //        Console.WriteLine($"Sheet: {sheetName}, SimpleDiff: {simpleDiff}");
                //        return;
                //    }
                //    continue;
                //}
                diffConfig.SrcSheetIndex = parentSheetNames.IndexOf(sheetName);
                Console.WriteLine($"Diffing Base and Mine");
                var mineDiff = ExcelSheet.Diff(ExcelWorkbook.GetSheet(parentFile, sheetName, readConfig)
                    , ExcelWorkbook.GetSheet(mineFile, sheetName, readConfig), diffConfig);
                var mineSummary = mineDiff.CreateSummary();
                Console.WriteLine($"Diffing Base and Other");
                var otherDiff = ExcelSheet.Diff(ExcelWorkbook.GetSheet(parentFile, sheetName, readConfig)
                    , ExcelWorkbook.GetSheet(othersFile, sheetName, readConfig), diffConfig);
                var otherSummary = otherDiff.CreateSummary();
                var diffSummories = new DiffSummories(sheetName, otherSummary, othersWorkbook, mineSummary, mineWorkbook);
                var confliced_cells = otherSummary.ModifiedCells.Where(x =>
                    mineSummary.ModifiedCells.Select(c => new Point(c.SrcCell.OriginalColumnIndex, c.SrcCell.OriginalRowIndex)).Contains(new Point(x.SrcCell.OriginalColumnIndex, x.SrcCell.OriginalRowIndex)));
                if (confliced_cells.Any())
                {
                    var conflictedCellsStr = string.Join(", ",
                        confliced_cells.Select(x => $"(Col: {x.SrcCell.OriginalColumnIndex}, Row:{x.SrcCell.OriginalRowIndex}) BaseValue:{x.SrcCell.Value}  OtherValue: {x.DstCell.Value};"));
                    Console.WriteLine($"Conflicted!! Sheet: {sheetName}, Confliced Cells: {conflictedCellsStr}");
                    return;
                }
                var inserted_column_indexes = otherSummary.InsertedColumnIndexes.Where(x =>
                    mineSummary.InsertedColumnIndexes.Contains(x));
                if (inserted_column_indexes.Any())
                {
                    Console.WriteLine($"Conflicted!! Sheet: {sheetName}, Inserted Column Indexes: {string.Join(", ", inserted_column_indexes)}");
                    return;
                }
                var inserted_row_indexes = otherSummary.InsertedRowIndexes.Where(x =>
                   mineSummary.InsertedRowIndexes.Contains(x));
                if (inserted_row_indexes.Any())
                {
                    Console.WriteLine($"Conflicted!! Sheet: {sheetName}, Inserted Row Indexes: {string.Join(", ", inserted_row_indexes)}");
                    return;
                }
                Console.WriteLine($"Merging Sheet: {sheetName}");
                int actual_inserted_column_count = 0;
                foreach (var cell in otherSummary.ModifiedCells)
                {
                    var srcCell = othersWorkbook.GetCell(sheetName,
                        cell.DstCell.OriginalRowIndex, cell.DstCell.OriginalColumnIndex);
                    var targetCell = diffSummories.GetMineCellFromOtherCell(srcCell);
                    ExcelUtility.CopyCellValue(srcCell, targetCell);
                }
                foreach (var idx in otherSummary.InsertedColumnIndexes)
                {
                    var column = othersSheet.GetColumn(idx);
                    int target_idx = idx + actual_inserted_column_count;
                    int inserted_row_count = 0;
                    mineWorkbook.InsertColumn(sheetName, target_idx, column, mineSummary.InsertedRowIndexes, (cell) =>
                    {
                        if (mineSummary.InsertedRowIndexes.Contains(cell.RowIndex))
                        {
                            inserted_row_count++;
                            //ExcelUtility.CopyCellValue(null, cell);
                        }
                        else
                        {
                            var srcCell = othersWorkbook.GetCell(sheetName, cell.RowIndex - inserted_row_count, idx);
                            //ExcelUtility.CopyCellValue(srcCell, cell);
                        }
                    });
                    actual_inserted_column_count++;
                }
                int actual_inserted_row_count = 0;
                foreach (var idx in otherSummary.InsertedRowIndexes)
                {
                    var row = othersSheet.Rows[idx]; int target_idx = idx + actual_inserted_row_count;
                    mineWorkbook.InsertRow(sheetName, target_idx, row, (cell, col_idx) =>
                    {
                        var srcCell = othersWorkbook.GetCell(sheetName, row.Index, col_idx);
                        ExcelUtility.CopyCellValue(srcCell, cell);
                    });
                    actual_inserted_row_count++;
                }
                mineWorkbook.Save();
            }
        }
    }
}
