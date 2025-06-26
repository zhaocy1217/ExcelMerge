using System.Collections.Generic;
using System.Linq;
using NetDiff;

namespace ExcelMerge
{
    public class ExcelSheetDiff
    {
        public SortedDictionary<int, ExcelColumnStatus> Columns { get; private set; }
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }

        public ExcelSheetDiff()
        {
            Columns = new SortedDictionary<int, ExcelColumnStatus>();
            Rows = new SortedDictionary<int, ExcelRowDiff>();
        }

        public ExcelColumnStatus CreateColumn(int index, ExcelColumnStatus status)
        {
            Columns.Add(index, status);
            return status;
        }
        public void SetColumnStatus(int index, ExcelColumnStatus status)
        {
            Columns[index] = status;
        }

        public ExcelRowDiff CreateRow()
        {
            var row = new ExcelRowDiff(Rows.Any() ? Rows.Keys.Last() + 1 : 0);
            Rows.Add(row.Index, row);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            var AllModifiedCells = new HashSet<ExcelCellDiff>();
            foreach (var row in Rows)
            {
                if (row.Value.IsAdded())
                    addedRowCount++;
                else if (row.Value.IsRemoved())
                    removedRowCount++;
                else if (row.Value.IsModified())
                {
                    modifiedRowCount++;
                    AllModifiedCells.UnionWith(row.Value.BaseSrcModifiedCells);
                }
                modifiedCellCount += row.Value.ModifiedCellCount;
            }

            return new ExcelSheetDiffSummary
            {
                Diff = this,
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
                InsertedColumnIndexes = Columns.Where(c => c.Value == ExcelColumnStatus.Inserted).Select(c => c.Key).ToList(),
                RemovedColumnIndexes = Columns.Where(c => c.Value == ExcelColumnStatus.Deleted).Select(c => c.Key).ToList(),
                InsertedRowIndexes = Rows.Where(r => r.Value.IsAdded()).Select(r => r.Key).ToList(),
                RemovedRowIndexes = Rows.Where(r => r.Value.IsRemoved()).Select(r => r.Key).ToList(),
                ModifiedRowIndexes = Rows.Where(r => r.Value.IsModified()).Select(r => r.Key).ToList(),
                ModifiedCells = AllModifiedCells,
            };
        }
    }
}
