using NPOI.SS.UserModel;
using TableRW.Read.NpoiEx;
using TableRW.Utils.Ex;

namespace TableRW.Read;

public static class SheetEx {
    public static List<TEntity> ReadToList<TEntity>(
        this ISheet sheet, int headerRow
    ) where TEntity : new() {
        if (sheet == null) { throw new ArgumentNullException(nameof(sheet)); }
        if (sheet.LastRowNum == 0) { throw new ArgumentNullException("sheet.LastRowNum == 0"); }

        if (CacheReadFn<TEntity>.FnUseHeader is {} fn) {
            return fn(sheet);
        }

        var header = GetHeader();
        if (header.Count == 0) {
            throw new InvalidOperationException("No data read: The number of column headers is 0");
        }

        var (iCol, m0) = header[0];
        var reader = new SheetReader<TEntity>();
        reader.SetStart(headerRow + 1, iCol);
        reader.AddColumn(m0);

        foreach (var (i, m) in header.Skip(1)) {
            if (i - iCol > 1) {
                reader.AddSkipColumn(i - iCol - 1);
            }
            reader.AddColumn(m);
            iCol = i;
        }

        var readLmd = reader.Lambda();
        CacheReadFn<TEntity>.FnUseHeader = fn = readLmd.Compile();
        return fn(sheet);

        List<(int i, MemberInfo member)> GetHeader() {
            var t_entity = typeof(TEntity);
            var props = t_entity.GetProperties().Where(p => p.CanWrite)
                .Concat<MemberInfo>(t_entity.GetFields().Where(f => !f.IsInitOnly))
                .Where(m => m.HasAttribute<IgnoreReadAttribute>() == false)
                .ToDictionary(m => m.Name);

            var row = sheet.GetRow(headerRow);
            return Enumerable.Range(0, row.LastCellNum)
                .Select(i => (i, text: row.Cells[i].StringCellValue))
                .Select((t) => (t.i, member: props.GetValueOr(t.text, null!)))
                .Where(t => t.member != null)
                .ToList();
        }
    }

    public static List<TEntity> ReadToList<TEntity>(
        this ISheet sheet,
        int cacheKey,
        Func<SheetReader<TEntity>, Func<ISheet, List<TEntity>>> buildRead
    ) where TEntity : new() {
        if (CacheReadFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildRead(new());
        }

        return fn(sheet);
    }

    public static List<TEntity> ReadToList<TEntity, TData>(
        this ISheet sheet,
        int cacheKey,
        Func<SheetReader<TEntity, TData>, Func<ISheet, List<TEntity>>> buildRead
    ) where TEntity : new() {
        if (CacheReadFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildRead(new());
        }

        return fn(sheet);
    }

}

static class CacheReadFn<T> {
    internal static Func<ISheet, List<T>>? FnUseHeader;

    internal static Dictionary<int, Func<ISheet, List<T>>> DicFn = new();
}
