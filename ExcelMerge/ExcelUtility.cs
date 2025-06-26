using NPOI.HSSF.UserModel;
using NPOI.POIFS.Common;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.Util;

namespace ExcelMerge
{
    public class ExcelUtility
    {
        public static void InsertColumn(ISheet sheet, int columnIndex, List<int> newInsertedRows, Action<ICell> onCellCreatAct)
        {
            // Get the maximum number of rows to process
            int lastRowNum = sheet.LastRowNum;

            // Process each row from bottom to top
            for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                // Get the last cell number in this row
                int lastCellNum = row.LastCellNum;
                int toBreakCount = 0;
                // Shift cells from right to left, starting from the rightmost cell
                for (int cellIndex = lastCellNum; cellIndex > columnIndex; cellIndex--)
                {
                    ICell sourceCell = row.GetCell(cellIndex - 1);
                    ICell targetCell = row.CreateCell(cellIndex, CellType.Blank);
                    targetCell.RemoveCellComment();
                    targetCell.RemoveFormula();
                    targetCell.RemoveHyperlink();
                    if (sourceCell != null)
                    {
                        CopyCellValue(sourceCell, targetCell);
                    }
                    toBreakCount++;
                }

                // Clear the inserted column cell
                ICell insertedCell = row.GetCell(columnIndex);
                if (insertedCell != null && onCellCreatAct != null)
                {
                    //onCellCreatAct(insertedCell);
                }
            }
        }
        public static void CopyCellComment(ICell sourceCell, ICell destCell)
        {
            IComment existingComment = sourceCell.CellComment;
            string commentText = existingComment.String.String;
            string commentAuthor = existingComment.Author;
            IClientAnchor existingAnchor = existingComment.ClientAnchor;
            IDrawing drawing = sourceCell.Sheet.CreateDrawingPatriarch();
            var colDiff = destCell.ColumnIndex - sourceCell.ColumnIndex;
            var rowDiff = destCell.RowIndex - sourceCell.RowIndex;
            IClientAnchor newAnchor = new XSSFClientAnchor(existingAnchor.Dx1, existingAnchor.Dy1, existingAnchor.Dx2, existingAnchor.Dy2,
                    existingAnchor.Col1 + colDiff, existingAnchor.Row1 + rowDiff, existingAnchor.Col2 + colDiff, existingAnchor.Row2 + rowDiff);
            newAnchor.AnchorType = existingAnchor.AnchorType;
            IComment newComment = drawing.CreateCellComment(newAnchor);
            newComment.String = existingComment.String;
            newComment.Author = commentAuthor;
            newComment.Row = destCell.RowIndex;
            newComment.Column = destCell.ColumnIndex;
            newComment.Visible = true;//existingComment.Visible;
            newComment.Address = new CellAddress(destCell.RowIndex, destCell.ColumnIndex);
            destCell.CellComment = newComment;
            sourceCell.CellComment = null;

        }
        private static void SetComment(ICell sourceCell, ICell targetCell)
        {
            if (sourceCell != null && sourceCell.CellComment != null)
            {
                CopyCellComment(sourceCell, targetCell);
                targetCell.CellStyle = sourceCell.CellStyle;
                var dataValidations = sourceCell.Sheet.GetDataValidations();
                if (dataValidations != null)
                {
                    foreach (var dataValidation in dataValidations)
                    {
                    }
                }
            }
        }
        public static void CopyCellValue(ICell sourceCell, ICell targetCell)
        {
            var cellType = sourceCell == null ? CellType.Blank : sourceCell.CellType;
            var oldCell = sourceCell;
            var newCell = targetCell;
            if (oldCell.CellStyle != null)
            {
                // apply style from old cell to new cell 
                newCell.CellStyle = oldCell.CellStyle;
            }

            // If there is a cell comment, copy
            if (oldCell.CellComment != null)
            {
                var s = oldCell == newCell;
                targetCell.Sheet.CopyComment(oldCell, newCell);
            }

            // If there is a cell hyperlink, copy
            if (oldCell.Hyperlink != null)
            {
                newCell.Hyperlink = oldCell.Hyperlink;
            }

            // Set the cell data type
            newCell.SetCellType(cellType);

            // Set the cell data value
            switch (cellType)
            {
                case CellType.Blank:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(oldCell))
                        newCell.SetCellValue(oldCell.DateCellValue.Value);
                    else
                        newCell.SetCellValue(oldCell.NumericCellValue);
                    break;
                case CellType.String:
                    newCell.SetCellValue(oldCell.RichStringCellValue);
                    break;
            }
        }
        public static object GetCellValue(ICell cell)
        {
            if (cell == null)
                return null;

            return GetCellValue(cell, cell.CellType);
        }

        private static object GetCellValue(ICell cell, CellType type)
        {
            if (cell != null)
            {
                switch (type)
                {
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            return cell.DateCellValue;
                        }
                        else
                        {
                            return cell.NumericCellValue;
                        }
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue;
                    case CellType.Formula:
                        return GetCellValue(cell, cell.CachedFormulaResultType);
                }
            }

            return string.Empty;
        }

        public static string GetCellStringValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;

            return GetCellValue(cell).ToString();
        }

        public static void CreateWorkbook(string path, ExcelWorkbookType workbookType)
        {
            if (!ValidateExtension(path, workbookType))
                throw new ArgumentException("The specified Excel type and path extension do not match.");

            var workbook = CreateWorkbook(workbookType);
            var sheet = workbook.CreateSheet();

            using (var fileStream = new FileStream(path, FileMode.Create))
            {
                workbook.Write(fileStream);
            }
        }
        private static IWorkbook CreateWorkbook(ExcelWorkbookType workbookType)
        {
            switch (workbookType)
            {
                case ExcelWorkbookType.XLS: return new HSSFWorkbook() as IWorkbook;
                case ExcelWorkbookType.XLSX: return new XSSFWorkbook() as IWorkbook;
                default: break;
            }

            throw new ArgumentException("The specified excel type is not supported instantiating.");
        }

        private static bool ValidateExtension(string path, ExcelWorkbookType workbookType)
        {
            switch (workbookType)
            {
                case ExcelWorkbookType.XLS: return Path.GetExtension(path) == ".xls";
                case ExcelWorkbookType.XLSX: return Path.GetExtension(path) == ".xlsx";
                default: break;
            }

            return false;
        }

        public static ExcelWorkbookType GetWorkbookType(string path)
        {
            var extension = Path.GetExtension(path);
            switch (extension)
            {
                case ".xls": return ExcelWorkbookType.XLS;
                case ".xlsx": return ExcelWorkbookType.XLSX;
                default: break;
            }

            return ExcelWorkbookType.None;
        }

        public static ExcelWorkbookType GetWorkboolTypeStrict(string path)
        {
            var type = GetWorkbookType(path);

            if (type == ExcelWorkbookType.None)
            {
                if (IsXLS(path))
                    type = ExcelWorkbookType.XLS;
                else if (IsXLSX(path))
                    type = ExcelWorkbookType.XLSX;
            }

            return type;
        }

        public static bool IsXLS(string path)
        {
            using (var inputStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                if (POIFSFileSystem.HasPOIFSHeader(inputStream))
                {
                    return true;
                }
            }
            return false;
        }
        public static bool HasOOXMLHeader(Stream inp)
        {
            // We want to peek at the first 4 bytes
            //inp.mark(4);

            byte[] header = new byte[4];
            int bytesRead = IOUtils.ReadFully(inp, header);

            // Wind back those 4 bytes
            if (inp is PushbackStream pin)
            {
                pin.Position = pin.Position - 4;
                //pin.unread(header, 0, bytesRead);
            }
            else
            {
                inp.Position = 0;
            }

            // Did it match the ooxml zip signature?
            return (
                bytesRead == 4 &&
                header[0] == POIFSConstants.OOXML_FILE_HEADER[0] &&
                header[1] == POIFSConstants.OOXML_FILE_HEADER[1] &&
                header[2] == POIFSConstants.OOXML_FILE_HEADER[2] &&
                header[3] == POIFSConstants.OOXML_FILE_HEADER[3]
            );
        }
        public static bool IsXLSX(string path)
        {
            using (var inputStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                if (HasOOXMLHeader(inputStream))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
