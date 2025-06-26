using System.Collections.Generic;
using NetDiff;

namespace ExcelMerge
{
    public struct ExcelSheetDiffSummary
    {
        public ExcelSheetDiff Diff { get; set; }
        public int ModifiedCellCount { get; set; }
        public int AddedRowCount { get; set; }
        public int RemovedRowCount { get; set; }
        public int ModifiedRowCount { get; set; }
        public List<int> InsertedColumnIndexes { get; set; }
        public List<int> RemovedColumnIndexes { get; set; }
        public List<int> InsertedRowIndexes { get; set; }
        public List<int> RemovedRowIndexes { get; set; }
        public List<int> ModifiedRowIndexes { get; set; }
        public HashSet<ExcelCellDiff> ModifiedCells { get; set; }
        public bool HasDiff { get { return ModifiedCellCount + AddedRowCount + RemovedRowCount + ModifiedRowCount > 0; } }
    }
}
