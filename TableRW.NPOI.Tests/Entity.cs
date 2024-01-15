namespace TableRW;

class RecordA {

    public RecordA() {
        Random random = new();
        _IgnoreWriteValue = random.Next();
        this.ReadonlyField = _IgnoreWriteValue;
        this.IgnoreField = _IgnoreWriteValue;
        this.ReadonlyProperty = _IgnoreWriteValue;
        this.InitProperty = _IgnoreWriteValue;
        this.IgnoreProperty = _IgnoreWriteValue;

    }

    int _IgnoreWriteValue;

    public string? FieldStr = null;
    public int FieldInt = 0;

    public readonly int ReadonlyField;
    [IgnoreRead]
    public int IgnoreField;

    public string? Str { get; set; }
    public int StructInt { get; set; }
    public int? NullableInt { get; set; }

    public int ReadonlyProperty { get; }
    [IgnoreRead]
    public int InitProperty { get; init; }
    [IgnoreRead]
    public int IgnoreProperty { get; set; }

    public void TestIgnoreWrite() {
        Assert.Equal(_IgnoreWriteValue, ReadonlyField);
        Assert.Equal(_IgnoreWriteValue, IgnoreField);
        Assert.Equal(_IgnoreWriteValue, ReadonlyProperty);
        Assert.Equal(_IgnoreWriteValue, InitProperty);
        Assert.Equal(_IgnoreWriteValue, IgnoreProperty);
    }

}
