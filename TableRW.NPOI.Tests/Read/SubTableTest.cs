using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using TableRW.Read;

namespace TableRW.Read.NpoiEx.Tests;


public class SubTableTest {
    class DataA {
        public int Id { get; set; }
        public string? Name { get; set; }
        public List<SubB> SubList { get; set; } = new();
    }

    class SubB {
        public int Id { get; set; }
        public string? Text { get; set; }
        public List<SubC> SubCList { get; set; } = new();
    }

    class SubC {
        public int Id { get; set; }
        public int NumC { get; set; }
    }

    readonly ISheet _sheetLevel2;
    readonly ISheet _sheetLevel3;
    public SubTableTest() {
        var excel = new XSSFWorkbook();
        _sheetLevel2 = excel.CreateSheet("l2");
        _sheetLevel2.WriteCells((0, 0), [
            [ 1000, "AA", 1, "Saa1" ],
            [ 1000, "AA", 2, "Saa2" ],
            [ 1000, "AA", 3, "Saa3" ],

            [ 2000, "BB", 1, "Sbb" ],

            [ 3000, "CC", 1, "Sccc" ],
            [ 3000, "CC", 2, "Sccc" ],
        ]);

        _sheetLevel3 = excel.CreateSheet("l3");
        _sheetLevel3.WriteCells((0, 0), [
            [ 1000, "AA", 1, "Saa1" ],

            [ 1000, "AA", 2, "Saa2", 2001, 201 ],
            [ 1000, "AA", 2, "Saa2", 2002, 202 ],
            [ 1000, "AA", 2, "Saa2", 2003, 203 ],

            [ 1000, "AA", 3, "Saa3", 3001, 301 ],
            [ 1000, "AA", 3, "Saa3", 3002, 302 ],

            [ 2000, "BB", 1, "Sbb" ],

            [ 3000, "CC", 1, "Scc1" ],
            [ 3000, "CC", 2, "Scc2" ],
        ]);

    }

    void CheckRead_SheetLevel2(List<DataA> list) {
        Assert.Equal(3, list.Count);

        Assert.Equal(1000, list[0].Id);
        Assert.Equal("AA", list[0].Name);
        Assert.Equal(3, list[0].SubList.Count);
        Assert.Equal(1, list[0].SubList[0].Id);
        Assert.Equal("Saa1", list[0].SubList[0].Text);

        Assert.Equal(2000, list[1].Id);
        Assert.Equal("BB", list[1].Name);
        Assert.Equal(1, list[1].SubList.Count);

        Assert.Equal(3000, list[2].Id);
        Assert.Equal("CC", list[2].Name);
        Assert.Equal(2, list[2].SubList.Count);
    }

    static void CheckRead_SheetLevel3(List<DataA> list) {
        Assert.Equal(3, list.Count);

        Assert.Equal(1000, list[0].Id);
        Assert.Equal("AA", list[0].Name);

        Assert.Equal(3, list[0].SubList.Count);
        Assert.Equal(1, list[0].SubList[0].Id);
        Assert.Equal("Saa1", list[0].SubList[0].Text);
        Assert.Equal(0, list[0].SubList[0].SubCList.Count);
        Assert.Equal(2, list[0].SubList[2].SubCList.Count);

        var subcList2 = list[0].SubList[1].SubCList;
        Assert.Equal(3, subcList2.Count);
        Assert.Equal(2002, subcList2[1].Id);
        Assert.Equal(202, subcList2[1].NumC);

        var subcList3 = list[0].SubList[2].SubCList;
        Assert.Equal(2, subcList3.Count);
        Assert.Equal(3002, subcList3[1].Id);
        Assert.Equal(302, subcList3[1].NumC);

        Assert.Equal(2000, list[1].Id);
        Assert.Equal("BB", list[1].Name);
        Assert.Equal(1, list[1].SubList.Count);
        Assert.Equal(0, list[1].SubList[0].SubCList.Count);

        Assert.Equal(3000, list[2].Id);
        Assert.Equal("CC", list[2].Name);
        Assert.Equal(2, list[2].SubList.Count);
        Assert.Equal(0, list[2].SubList[0].SubCList.Count);
    }

    [Fact]
    public void Level2_AddSubTable() {
        var reader = new SheetReader<DataA>()
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddSubTable(e => e.SubList, (s, e) => s(e.Id, e.Text));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(_sheetLevel2);
        CheckRead_SheetLevel2(list);
    }

    [Fact]
    public void Level2_AddSubTable_WithData() {
        var reader = new SheetReader<DataA, (string name, int count)>()
            .InitData(src => (src.SheetName, src.LastRowNum))
            .AddActionRead(it => {
                Assert.Equal(_sheetLevel2.SheetName, it.Data.name);
                Assert.Equal(_sheetLevel2.LastRowNum, it.Data.count);
            })
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddSubTable(e => e.SubList, sub => sub
                .AddColumns((s, e) => s(e.Id, e.Text))
                .AddActionRead(it => {
                    DataA dataA = it.Parent.Entity;
                    Assert.NotEqual(0, dataA.Id);

                    Assert.Equal(_sheetLevel2.SheetName, it.Data.name);
                    Assert.Equal(_sheetLevel2.LastRowNum, it.Data.count);
                }));

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(_sheetLevel2);
        CheckRead_SheetLevel2(list);
    }

    [Fact]
    public void Level3_AddSubTable() {
        var reader = new SheetReader<DataA>()
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddSubTable(e => e.SubList, sub => sub
                .AddColumns((s, e) => s(s.RowKey(e.Id), e.Text))
                .AddSubTable(e => e.SubCList, (s, e) => s(s.RowKey(e.Id), e.NumC))
            );

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(_sheetLevel3);
        CheckRead_SheetLevel3(list);
    }

    [Fact]
    public void Level3_AddSubTable_Action() {
        var reader = new SheetReader<DataA>()
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddSubTable(e => e.SubList, sub => sub
                .AddColumns((s, e) => s(s.RowKey(e.Id), e.Text))
                .AddSubTable(e => e.SubCList, sub => sub
                    .AddColumns((s, e) => s(s.RowKey(e.Id), e.NumC)))
            );

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(_sheetLevel3);
        CheckRead_SheetLevel3(list);
    }

    [Fact]
    public void Level3_AddSubTable_WithData_Member() {
        var reader = new SheetReader<DataA, (string name, int count)>()
            .InitData(src => (src.SheetName, src.LastRowNum))
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddSubTable(e => e.SubList, sub => sub
                .AddColumns((s, e) => s(s.RowKey(e.Id), e.Text))
                .AddSubTable(e => e.SubCList, (s, e) => s(s.RowKey(e.Id), e.NumC))
            );

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(_sheetLevel3);
        CheckRead_SheetLevel3(list);
    }

    [Fact]
    public void Level3_AddSubTable_WithData_Action() {
        var reader = new SheetReader<DataA, (string name, int count)>()
            .InitData(src => (src.SheetName, src.LastRowNum))
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddActionRead(it => {
                Assert.Equal(_sheetLevel3.SheetName, it.Data.name);
                Assert.Equal(_sheetLevel3.LastRowNum, it.Data.count);
            })
            .AddSubTable(e => e.SubList, sub => sub
                .AddColumns((s, e) => s(s.RowKey(e.Id), e.Text))
                .AddActionRead(it => {
                    DataA dataA = it.Parent.Entity;
                    Assert.NotEqual(0, dataA.Id);

                    Assert.Equal(_sheetLevel3.SheetName, it.Data.name);
                    Assert.Equal(_sheetLevel3.LastRowNum, it.Data.count);
                })
                .AddSubTable(e => e.SubCList, sub => sub
                    .AddColumns((s, e) => s(s.RowKey(e.Id), e.NumC))
                    .AddActionRead(it => {
                        SubB subB = it.Parent.Entity;
                        Assert.NotEqual(0, subB.Id);

                        DataA dataA = it.Parent.Parent.Entity;
                        Assert.NotEqual(0, dataA.Id);

                        Assert.Equal(_sheetLevel3.SheetName, it.Data.name);
                        Assert.Equal(_sheetLevel3.LastRowNum, it.Data.count);
                    }))
            );

        var readLmd = reader.Lambda();
        var readFn = readLmd.Compile();
        var list = readFn(_sheetLevel3);
        CheckRead_SheetLevel3(list);
    }

    [Fact]
    public void Level3_AddSubTable_Event() {
        var reader = new SheetReader<DataA, List<int>>()
            .InitData(src => new(40))
            .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name))
            .AddSubTable(e => e.SubList, sub => sub
                .AddColumns((s, e) => s(s.RowKey(e.Id), e.Text))
                .AddSubTable(e => e.SubCList, sub => sub
                    .OnStartReadingTable(it => {
                        it.Data.Add(it.Parent.Parent.Entity.Id);
                        return true;
                    })
                    .OnStartReadingRow(it => {
                        it.Data.Add(it.Parent.Parent.Entity.Id);
                    })
                    .AddColumns((s, e) => s(s.RowKey(e.Id), e.NumC))
                    .OnEndReadingRow(it => {
                        it.Data.Add(it.Entity.Id);
                    })
                    .OnEndReadingTable(it => {
                        it.Data.Add(it.Entity.Id);
                    }))
            );

        var readLmd = reader.Lambda(f => f.ReturnData());
        var readFn = readLmd.Compile();
        var (list, data) = readFn(_sheetLevel3);
        CheckRead_SheetLevel3(list);
        Assert.Equal(data, [
            1000, // on start table
            1000, 2001,
            1000, 2002,
            1000, 2003,
            2003, // on end table
            1000,
            1000, 3001,
            1000, 3002,
            3002,
        ]);

    }

}
