using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelMerge;
using NetDiff;
using NPOI.HSSF.Record.PivotTable;
using NPOI.SS.UserModel;

namespace MergeDiff
{
    static class MergeDiffUtility
    {
        public static int GetHashCode(IEnumerable<string> iEnumrable)
        {
            var totalStr = "";
            foreach (var item in iEnumrable)
            {
                totalStr += item;
            }
            return totalStr.GetHashCode();
        }
        public static bool ContainsChinese(string input)
        {
            if (string.IsNullOrEmpty(input))
                return false;
            return Regex.IsMatch(input, @"[\u4e00-\u9fff]");
        }
    }
    class DiffSummories
    {
        Dictionary<Point, Point> _otherToBaseMap = new Dictionary<Point, Point>();
        Dictionary<Point, Point> _mineToBaseMap = new Dictionary<Point, Point>();
        ExcelSheetDiffSummary _otherDiff;
        ExcelSheetDiffSummary _mineDiff;
        ExcelWorkbook _othersWorkbook;
        ExcelWorkbook _mineWorkbook;
        string _sheetName;
        public DiffSummories(string sheetName, ExcelSheetDiffSummary otherDiff, ExcelWorkbook othersWorkbook, ExcelSheetDiffSummary mineDiff, ExcelWorkbook mineWorkbook)
        {
            this._otherDiff = otherDiff;
            this._mineDiff = mineDiff;
            this._othersWorkbook = othersWorkbook;
            this._mineWorkbook = mineWorkbook;
            this._sheetName = sheetName;
            foreach (var cell in otherDiff.ModifiedCells)
            {
                _otherToBaseMap.Add(new Point(cell.DstCell.OriginalRowIndex, cell.DstCell.OriginalColumnIndex),
                    new Point(cell.SrcCell.OriginalRowIndex, cell.SrcCell.OriginalColumnIndex));
            }
            foreach (var cell in mineDiff.ModifiedCells)
            {
                _mineToBaseMap.Add(new Point(cell.DstCell.OriginalRowIndex, cell.DstCell.OriginalColumnIndex),
                    new Point(cell.SrcCell.OriginalRowIndex, cell.SrcCell.OriginalColumnIndex));
            }
        }
        public ICell GetMineCellFromOtherCell(ICell otherCell)
        {
            return GetMineCellFromOtherCell(new Point(otherCell.RowIndex, otherCell.ColumnIndex));
        }
        private int GetIndexOfSrc(ExcelRowDiff row)
        {
            if (!row.IsAdded() && !row.IsRemoved())
            {
                if (row.Cells.Any())
                {
                    var firstCell = row.Cells.First().Value;
                    if (firstCell != null && firstCell.SrcCell != null)
                    {
                        return firstCell.SrcCell.OriginalRowIndex;
                    }
                    else
                    {
                        Console.WriteLine($"Row: {row.Index}, FirstCell: {firstCell.SrcCell.OriginalRowIndex}, {firstCell.SrcCell.OriginalColumnIndex} is null");
                    }
                }
            }
            return -1;
        }
        public bool FindSrcCell(Point point, ExcelSheetDiff diff, out ICell targetCell)
        {
            var rowList = diff.Rows.Values.ToList();
            var idx = rowList.Select(GetIndexOfSrc).ToList().BinarySearch(point.X);
            if (idx != -1)
            {
                var cells = rowList[idx].Cells;
                var cell_dict_pair = cells.Where(c => c.Value.SrcCell.OriginalColumnIndex == point.Y).FirstOrDefault();
                var cell = cell_dict_pair.Value;
                if (cell != null)
                {
                    targetCell = _mineWorkbook.GetCell(this._sheetName,
                        cell.DstCell.OriginalRowIndex, cell.DstCell.OriginalColumnIndex);
                    return true;
                }
            }
            targetCell = null;
            return false;
        }
        public ICell GetMineCellFromOtherCell(Point otherCell)
        {
            if (_otherToBaseMap.TryGetValue(otherCell, out var baseCell)
                && FindSrcCell(baseCell, _mineDiff.Diff, out var mineCell))
            {
                return mineCell;
            }
            return null;
        }
    }
}
