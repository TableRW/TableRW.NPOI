
using NPOI.XSSF.UserModel;

namespace TableRW.Read.NpoiEx.Tests;

public class DataTableReaderTest {


    [Fact]
    public void AddColumns() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
            ["A1", 21, "C1", 41, null],
            ["A2", 22, "C2", 42, 52],
        ]);


        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldStr, s.Skip(2), e.StructInt, e.NullableInt));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);

        Assert.Equal(2, list.Count);
        Assert.Equal("A1", list[0].FieldStr);
        Assert.Equal(41, list[0].StructInt);
        Assert.Null(list[0].NullableInt);
        Assert.Equal("A2", list[1].FieldStr);
        Assert.Equal(42, list[1].StructInt);
        Assert.Equal(52, list[1].NullableInt);
    }


    [Fact]
    public void AddActionRead_Action() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
            ["30"],
            ["ss"],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddActionRead(it => it.Entity.Str = it.Src.GetRow(it.iRow).GetCell(0).StringCellValue + "-")
            .AddColumns((s, e) => s(e.FieldStr));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Equal(2, list.Count);
        Assert.Equal("30-", list[0].Str);
        Assert.Equal("ss-", list[1].Str);
        Assert.Equal("30", list[0].FieldStr);
        Assert.Equal("ss", list[1].FieldStr);
    }

    [Fact]
    public void AddActionRead_SkipRow() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
            [ 1, "30" ],
            [ 2, "ss" ],
            [ 3, "33" ],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldInt))
            .AddActionRead(it => {
                var val = it.Src.GetRow(it.iRow).GetCell(1).StringCellValue;
                return it.SkipRow(!int.TryParse(val, out var _));
            })
            .AddColumns((s, e) => s(e.Str));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Equal(2, list.Count);
        Assert.Equal(1, list[0].FieldInt);
        Assert.Equal(3, list[1].FieldInt);
        Assert.Equal("30", list[0].Str);
        Assert.Equal("33", list[1].Str);
    }

    [Fact]
    public void AddActionRead_EndRow() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
            [ 1, "30" ],
            [ 2, "ss" ],
            [ 3, "33" ],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldInt))
            .AddActionRead(it => {
                var val = it.Src.GetRow(it.iRow).GetCell(1).StringCellValue;
                return it.EndRow(!int.TryParse(val, out var _));
            })
            .AddColumns((s, e) => s(e.Str));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Equal(3, list.Count);
        Assert.Equal(1, list[0].FieldInt);
        Assert.Equal(2, list[1].FieldInt);
        Assert.Equal(3, list[2].FieldInt);
        Assert.Equal("30", list[0].Str);
        Assert.Null(list[1].Str);
        Assert.Equal("33", list[2].Str);
    }

    [Fact]
    public void AddActionRead_EndTable() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
            [ 1, "30" ],
            [ 2, "ss" ],
            [ 3, "33" ],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldInt))
            .AddActionRead(it => {
                var val = it.Src.GetRow(it.iRow).GetCell(1).StringCellValue;
                return it.EndTable(!int.TryParse(val, out var _));
            })
            .AddColumns((s, e) => s(e.Str));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Single(list);
        Assert.Equal(1, list[0].FieldInt);
        Assert.Equal("30", list[0].Str);
    }

    [Fact]
    public void AddColumnRead_Action() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
            [ 1, "30" ],
            [ 2, "ss" ],
        ]);
        var reader = new SheetReader<RecordA>()
            .AddColumnRead((int val) => it => it.Entity.StructInt = val + 10)
            .AddColumns((s, e) => s(e.FieldStr));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Equal(2, list.Count);
        Assert.Equal(11, list[0].StructInt);
        Assert.Equal(12, list[1].StructInt);
        Assert.Equal("30", list[0].FieldStr);
        Assert.Equal("ss", list[1].FieldStr);
    }

    [Fact]
    public void AddColumnRead_SkipRow() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
                [ 11, "21", "31" ],
                [ 12, "ss", "32" ],
                [ 13, "23", "33" ],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldInt))
            .AddColumnRead((string val) => it => {
                if (int.TryParse(val, out var cellVal)) {
                    it.Entity.StructInt = cellVal;
                    return null;
                }
                return it.SkipRow();
            })
            .AddColumns((s, e) => s(e.Str));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Equal(2, list.Count);
        Assert.Equal(11, list[0].FieldInt);
        Assert.Equal(13, list[1].FieldInt);
        Assert.Equal(21, list[0].StructInt);
        Assert.Equal(23, list[1].StructInt);
        Assert.Equal("31", list[0].Str);
        Assert.Equal("33", list[1].Str);
    }

    [Fact]
    public void AddColumnRead_EndRow() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
                [ 11, "21", "31" ],
                [ 12, "ss", "32" ],
                [ 13, "23", "33" ],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldInt))
            .AddColumnRead((string val) => it => {
                if (int.TryParse(val, out var intVal)) {
                    it.Entity.StructInt = intVal;
                    return null;
                }
                return it.EndRow();
            })
            .AddColumns((s, e) => s(e.Str));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Equal(3, list.Count);
        Assert.Equal(11, list[0].FieldInt);
        Assert.Equal(12, list[1].FieldInt);
        Assert.Equal(13, list[2].FieldInt);
        Assert.Equal(21, list[0].StructInt);
        Assert.Equal(00, list[1].StructInt);
        Assert.Equal(23, list[2].StructInt);
        Assert.Equal("31", list[0].Str);
        Assert.Null(list[1].Str);
        Assert.Equal("33", list[2].Str);
    }

    [Fact]
    public void AddColumnRead_EndTable() {
        var excel = new XSSFWorkbook();
        var sheet1 = excel.CreateSheet("sheet1");
        sheet1.WriteCells(start: (0, 0), [
                [ 11, "21", "31" ],
                [ 12, "ss", "32" ],
                [ 13, "23", "33" ],
        ]);

        var reader = new SheetReader<RecordA>()
            .AddColumns((s, e) => s(e.FieldInt))
            .AddColumnRead((string val) => it => {
                if (int.TryParse(val, out var cellVal)) {
                    it.Entity.StructInt = cellVal;
                    return null;
                }
                return it.EndTable();
            })
            .AddColumns((s, e) => s(e.Str));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(sheet1);
        Assert.Single(list);
        Assert.Equal(11, list[0].FieldInt);
        Assert.Equal(21, list[0].StructInt);
        Assert.Equal("31", list[0].Str);
    }

}
