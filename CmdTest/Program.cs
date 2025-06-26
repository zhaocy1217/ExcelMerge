using ExcelMerge;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace CmdTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            (string parentFile, string othersFile, string mineFile) = (args[0], args[1], args[2]);
            var readConfig = new ExcelSheetReadConfig() { TrimLastBlankColumns = true, TrimLastBlankRows = true };
            var parentWorkbook = ExcelWorkbook.Create(parentFile, readConfig);
            var sheet = parentWorkbook.RawWorkbook.GetSheet("Sheet1");
            ISheet sheet1 = parentWorkbook.RawWorkbook.GetSheetAt(0); // Get the first sheet
            // Iterating through cells to find comments
            
            foreach (IRow row1 in sheet1)
            {
                foreach (ICell cell in row1)
                {
                    if (cell.CellComment != null)
                    {
                        IComment comment = cell.CellComment;
                        Console.WriteLine($"Cell {cell.Address} has comment by '{comment.Author}': {comment.String.String}");

                        // Example: Remove a specific comment
                        // if (comment.Author == "Some Author")
                        // {
                        //    cell.RemoveCellComment();
                        // }
                    }
                }
            }

            var row = sheet.GetRow(0);
            var cellIndex = 2;
            var sourceCell = row.Cells[cellIndex];
            //if (sourceCell != null)
            //{
            //    ICell targetCell = row.CreateCell(cellIndex + 1, sourceCell.CellType);
            //    //CellUtil.CopyCell(row, sourceCell.ColumnIndex, targetCell.ColumnIndex);
            //    ExcelUtility.CopyCellValue(sourceCell, targetCell);
            //}
            parentWorkbook.Save();
        }
    }
}
