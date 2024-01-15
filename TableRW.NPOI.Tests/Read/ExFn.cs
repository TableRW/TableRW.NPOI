
using NPOI.SS.UserModel;

namespace TableRW.Read.NpoiEx.Tests;

static class ExcelEx {
    public static void WriteCells(this ISheet sheet, (int row, int col) start, object?[][] values) {
        for (int iRow = 0; iRow < values.Length; iRow++) {
            var row = sheet.CreateRow(start.row + iRow);
            for (int iCol = 0; iCol < values[iRow].Length; iCol++) {
                if (values[iRow][iCol] is var val && val == null) { continue; }

                var cell = row.CreateCell(start.col + iCol);
                (val switch {
                    string x => () => cell.SetCellValue(x),
                    double x => () => cell.SetCellValue(x),
                    int x => () => cell.SetCellValue(x),
                    bool x => () => cell.SetCellValue(x),
                    DateTime x => () => cell.SetCellValue(x),
                    _ => (Action)null!,
                })();

            }
        }

    }
}