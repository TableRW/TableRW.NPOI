using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TableRW.Write.NpoiEx.Tests;
public class EventTest {

    record EntityA(int Int1, int Int2, int? NullableInt, string Str1, DateTime Date);

    List<EntityA> DataList = new() {
        new(1, 11, 1000, "aaa", new(2021, 1, 11)),
        new(3, 33, null, "ccc", new(2022, 3, 13)),
        new(7, 77, 7000, "bbb", new(2023, 7, 17)),
    };

    readonly IWorkbook _excel;
    readonly ISheet _sheet;

    public EventTest() {
        _excel = new XSSFWorkbook();
        _sheet = _excel.CreateSheet("sheet1");
    }

    [Fact]
    public void StartWritingTable() {
        var (startRow, startCol) = (2, 2);
        var writer = new SheetWriter<EntityA, string>()
            .SetStart(startRow, startCol)
            .OnStartWritingTable(it => {
                Assert.Equal(2, it.iCol);
                it.Data = "data";
            })
            .AddColumn(it => it.Row.CreateCell(it.iCol).SetCellValue(it.Data))
            .AddColumns((s, e) => s(e.Int1, e.Str1));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_sheet, DataList);

        Assert.Equal(DataList.Count + startRow, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(startRow + i);
            var col = startCol;
            Assert.Equal("data", row.GetCell(col++).StringCellValue);
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
        }
    }

    [Fact]
    public void StartWritingRow() {
        var writer = new SheetWriter<EntityA>()
            .SetStart(0, 2)
            .OnStartWritingRow(it => it.Row.CreateCell(it.iCol - 1).SetCellValue(222))
            .AddColumns((s, e) => s(e.Int1, e.Int2, e.Str1));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 1;
            Assert.Equal(222, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Int2, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
        }
    }

    [Fact]
    public void EndWritingRow() {
        var writer = new SheetWriter<EntityA>()
            .AddColumns((s, e) => s(e.Int1, e.Int2, e.Date.Day))
            .OnEndWritingRow(it => it.Row.CreateCell(it.iCol + 1).SetCellValue(222));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Int2, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Date.Day, row.GetCell(col++).NumericCellValue);
            Assert.Equal(222, row.GetCell(col++).NumericCellValue);
        }
    }

    [Fact]
    public void EndWritingTable() {
        var writer = new SheetWriter<EntityA>()
            .AddColumns((s, e) => s(e.Int1, e.Int2, e.Date.Day))
            .OnEndWritingTable(it => it.Row.CreateCell(it.iCol + 1).SetCellValue(444));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Int2, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Date.Day, row.GetCell(col++).NumericCellValue);
        }
        Assert.Equal(444, _sheet.GetRow(DataList.Count - 1).Cells[3].NumericCellValue);
    }
}
