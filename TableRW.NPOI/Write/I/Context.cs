using NPOI.SS.UserModel;

namespace TableRW.Write.I.NpoiEx;
public class Context<TEntity>(ISheet src)
: I.Context<ISheet, IRow, TEntity>(src) { }

public class Context<TEntity, TData>(ISheet src)
: I.Context<ISheet, IRow, TEntity, TData>(src) { }

