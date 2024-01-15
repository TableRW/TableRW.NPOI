using NPOI.SS.UserModel;

namespace TableRW.Read.I.NpoiEx;

public class Context<TEntity>(ISheet src)
: I.Context<ISheet, TEntity>(src) { }

public class Context<TEntity, TData>(ISheet src)
: I.Context<ISheet, TEntity, TData>(src) { }
