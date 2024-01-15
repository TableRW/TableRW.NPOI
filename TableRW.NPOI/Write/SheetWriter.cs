
using TableRW.Write.I.NpoiEx;

namespace TableRW.Write.NpoiEx;

public class SheetWriter<TEntity> : SheetWriterImpl<Context<TEntity>> { }

public class SheetWriter<TEntity, TData> : SheetWriterImpl<Context<TEntity, TData>> { }
