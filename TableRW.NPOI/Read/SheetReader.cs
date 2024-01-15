
using TableRW.Read.I.NpoiEx;

namespace TableRW.Read.NpoiEx;

public class SheetReader<TEntity>
: SheetReaderImpl<Context<TEntity>> { }

public class SheetReader<TEntity, TData>
: SheetReaderImpl<Context<TEntity, TData>> { }
