
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TableRW.Write.NpoiEx.Tests;

public class SheetExTest {

    record EntityA(int Int1, int Int2, int? NullableInt, string Str1, DateTime Date);

    List<EntityA> DataList = new() {
        new(10, 11, 1000, "aaa", new(2023, 1, 1)),
        new(30, 33, 3000, "ccc", new(2023, 3, 3)),
        new(70, 77, 7000, "bbb", new(2023, 7, 7)),
    };

    readonly IWorkbook _excel;
    ISheet _sheet;

    public SheetExTest() {
        _excel = new XSSFWorkbook();
        _sheet = _excel.CreateSheet("sheet1");
    }


    [Fact]
    public void WriteFrom() {
        _sheet.WriteFrom(DataList, cacheKey: 0, writer => {
            writer.AddColumns((s, e) =>
                s(e.Int1, s.Skip(1), e.NullableInt, e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);
        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;

            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(null, row.GetCell(col++));
            Assert.Equal(e.NullableInt!.Value, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
            Assert.Equal(null, row.GetCell(col++));
        }
    }

    [Fact]
    public void WriteFrom_AnotherKey() {
        _sheet.WriteFrom(DataList, cacheKey: 1, writer => {
            writer.AddColumns((s, e) =>
                s(e.Int1, s.Skip(1), e.NullableInt, e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });
        _excel.RemoveSheetAt(0);
        _sheet = _excel.CreateSheet();

        // 使用另一种方式，和上面的缓存不同
        _sheet.WriteFrom(DataList, cacheKey: 2, writer => {
            writer.AddColumns((s, e) =>
                s(e.Int2, e.Int1, s.Skip(2), e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);
        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;

            Assert.Equal(e.Int2, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(null, row.GetCell(col++));
            Assert.Equal(null, row.GetCell(col++));
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
            Assert.Equal(null, row.GetCell(col++));
        }
    }

    [Fact]
    public void WriteFrom_WithData() {
        _sheet.WriteFrom<EntityA, string?>(DataList, cacheKey: 3, writer => {
            writer
                .InitData(src => src.SheetName)
                .AddColumn(it => it.Row.CreateCell(it.iCol).SetCellValue(it.Data))
                .AddColumns((s, e) => s(e.Int1, e.Int2, e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });

        Assert.Equal(DataList.Count, _sheet.LastRowNum + 1);
        for (var i = 0; i < DataList.Count; i++) {
            var e = DataList[i];
            var row = _sheet.GetRow(i);
            var col = 0;

            Assert.Equal(_sheet.SheetName, row.GetCell(col++).StringCellValue);
            Assert.Equal(e.Int1, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Int2, row.GetCell(col++).NumericCellValue);
            Assert.Equal(e.Str1, row.GetCell(col++).StringCellValue);
        }
    }

}
