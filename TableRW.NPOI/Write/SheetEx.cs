
using NPOI.SS.UserModel;

namespace TableRW.Write.NpoiEx;

public static class SheetEx {

    public static void WriteFrom<TEntity>(
        this ISheet tbl,
        IEnumerable<TEntity> enumerable,
        int cacheKey,
        Func<SheetWriter<TEntity>, Action<ISheet, IEnumerable<TEntity>>> buildWrite
    ) {
        if (CacheFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildWrite(new());
        }

        fn(tbl, enumerable);
    }

    public static void WriteFrom<TEntity, TData>(
        this ISheet tbl,
        IEnumerable<TEntity> enumerable,
        int cacheKey,
        Func<SheetWriter<TEntity, TData>, Action<ISheet, IEnumerable<TEntity>>> buildWrite
    ) {
        if (CacheFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildWrite(new());
        }

        fn(tbl, enumerable);
    }
}

static class CacheFn<T> {
    internal static Dictionary<int, Action<ISheet, IEnumerable<T>>> DicFn = new();
}
