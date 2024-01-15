
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TableRW.Write.NpoiEx.Tests;

public class WriteColumnTest {

    record EntityA(int Int1, int Int2, int? NullableInt, string Str1, DateTime Date);

    List<EntityA> DataList = new() {
        new(1, 11, 1000, "aaa", new(2023, 1, 1)),
        new(3, 33, null, "ccc", new(2023, 3, 3)),
        new(7, 77, 7000, "bbb", new(2023, 7, 7)),
    };


    readonly IWorkbook _excel;
    readonly ISheet _sheet;

    public WriteColumnTest() {
        _excel = new XSSFWorkbook();
        _sheet = _excel.CreateSheet("sheet1");
    }


    [Fact]
    public void AddColumns() {
        var writer = new SheetWriter<EntityA>()
            .AddColumns((s, e) => s(e.Int1, s.Skip(1), e.NullableInt, e.Str1));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);

        var iCol = 0;
        var row = _sheet.GetRow(0);
        Assert.Equal(DataList[0].Int1, row.GetCell(iCol++).NumericCellValue);
        Assert.Equal(null, row.GetCell(iCol++));
        Assert.Equal(DataList[0].NullableInt!.Value, row.GetCell(iCol++).NumericCellValue);
        Assert.Equal(DataList[0].Str1, row.GetCell(iCol++).StringCellValue);

        iCol = 0;
        row = _sheet.GetRow(1);
        Assert.Equal(DataList[1].Int1, row.GetCell(iCol++).NumericCellValue);
        Assert.Equal(null, row.GetCell(iCol++));
        Assert.Equal(CellType.Blank, row.GetCell(iCol++).CellType);
        Assert.Equal(DataList[1].Str1, row.GetCell(iCol++).StringCellValue);
    }


    [Fact]
    public void AddColumns_Compute() {
        var writer = new SheetWriter<EntityA>()
            .AddColumns((s, e) => s(
                e.Int1, e.Int1 + e.Int2, e.NullableInt + 1000,
                $"{e.Str1} -- {e.Date.Month}",
                "AA " + DateTime.Now.Month
            ));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);


        var col = 0;
        var e = DataList[0];
        var row = _sheet.GetRow(0);
        Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
        Assert.Equal(e.Int1 + e.Int2, row.GetCell(col++).NumericCellValue);
        Assert.Equal(e.NullableInt + 1000 ?? 0, row.GetCell(col++).NumericCellValue);
        Assert.Equal($"{e.Str1} -- {e.Date.Month}", row.GetCell(col++).StringCellValue);
        Assert.Equal("AA " + DateTime.Now.Month, row.GetCell(col++).StringCellValue);

        col = 0;
        e = DataList[1];
        row = _sheet.GetRow(1);
        Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
        Assert.Equal(e.Int1 + e.Int2, row.GetCell(col++).NumericCellValue);
        Assert.Equal(CellType.Blank, row.GetCell(col++).CellType);
        Assert.Equal($"{e.Str1} -- {e.Date.Month}", row.GetCell(col++).StringCellValue);
        Assert.Equal("AA " + DateTime.Now.Month, row.GetCell(col++).StringCellValue);

        col = 0;
        e = DataList[1];
        row = _sheet.GetRow(1);
        Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
        Assert.Equal(e.Int1 + e.Int2, row.GetCell(col++).NumericCellValue);
        Assert.Equal(e.NullableInt + 1000 ?? 0, row.GetCell(col++).NumericCellValue);
        Assert.Equal($"{e.Str1} -- {e.Date.Month}", row.GetCell(col++).StringCellValue);
        Assert.Equal("AA " + DateTime.Now.Month, row.GetCell(col++).StringCellValue);

    }


    [Fact]
    public void AddSkipColumn() {
        var writer = new SheetWriter<EntityA>()
            .AddSkipColumn(1)
            .AddColumns((s, e) => s(e.Int1))
            .AddSkipColumn(1)
            .AddColumns((s, e) => s(e.Str1))
            .AddSkipColumn(1)
            .AddAction(it => Assert.Equal(4, it.iCol))
            .AddAction(it => it.Row.CreateCell(it.iCol).SetCellValue("E"))
            .AddColumn(it => it.Row.CreateCell(it.iCol).SetCellValue(it.Entity.Str1));

        var writeLmd = writer.Lambda();
        var fn = writeLmd.Compile();
        fn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;
            Assert.Equal(null, row.GetCell(col++));
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(null, row.GetCell(col++));
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
            Assert.Equal("E", row.GetCell(col++).StringCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
        }

    }

    [Fact]
    public void AddActionWrite() {
        var writer = new SheetWriter<EntityA>()
            .AddAction(it => it.Row.CreateCell(it.iCol).SetCellValue(1111))
            .AddSkipColumn(2)
            .AddColumns((s, e) => s(e.Int1, e.Str1))
            .AddAction(it => it.Row.CreateCell(it.iCol).SetCellValue(it.Entity.Str1 + "-A"))
            .AddColumn(it => it.Row.CreateCell(it.iCol).SetCellValue(it.Entity.Str1))
            .AddAction(it => Assert.Equal(4, it.iCol));

        var writeLmd = writer.Lambda();
        var fn = writeLmd.Compile();
        fn(_sheet, DataList);

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;
            Assert.Equal(1111, row.GetCell(col++).NumericCellValue);
            Assert.Equal(null, row.GetCell(col++));
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Str1 + "-A", row.GetCell(col++).StringCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
        }
    }

    [Fact]
    public void SetStart() {
        var (startRow, startCol) = (2, 2);
        var writer = new SheetWriter<EntityA>()
            .SetStart(startRow, startCol)
            .AddColumns((s, e) => s(e.Int1, e.Str1));

        var writeLmd = writer.Lambda();
        var fn = writeLmd.Compile();
        fn(_sheet, DataList);

        Assert.Equal(DataList.Count + startRow, _sheet.LastRowNum + 1);

        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i + startRow);
            var col = startRow;
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
        }
    }


}
