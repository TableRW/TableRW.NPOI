using NPOI.SS.UserModel;
using E = System.Linq.Expressions.Expression;

namespace TableRW.Read.I.NpoiEx;

public class SheetReaderImpl<C> : TableReaderImpl<C> {

    static SheetReaderImpl() {
        ReadSource<ISheet>.Impl(
            ReadSrcValueByIndex,
            (src, iRow) => iRow > src.LastRowNum,
            (src, iRow, iCol) => iCol >= src.GetRow(iRow).LastCellNum
                || src.GetRow(iRow).GetCell(iCol).CellType == CellType.Blank);
    }

    public static Expression ReadSrcValueByIndex(Expression ctx, Type valueType)
        => ConvertSrcValue(GetSrcValueByIndex(ctx), valueType);

    public static Expression GetSrcValueByIndex(Expression ctx)
        => Utils.Expr.ExtractBody((ISource<ISheet> ctx) =>
            ctx.Src.GetRow(ctx.iRow).GetCell(ctx.iCol), ctx);

    public static Expression ConvertSrcValue(Expression cell, Type valueType) {
        if (Nullable.GetUnderlyingType(valueType) is { } vType) {
            // cell == null || cell.CellType == CellType.Blank
            // ? null
            // : Nullalbe( (cell.CellValue) )
            return E.Condition(
                E.OrElse(
                    E.Equal(cell, E.Constant(null)),
                    E.Equal(E.Property(cell, nameof(ICell.CellType)), E.Constant(CellType.Blank))),
                E.Constant(null, valueType),
                E.Convert(ConvertSrcValue(cell, vType), valueType));
        }
        if (valueType == typeof(DateTimeOffset)) {
            return E.Convert(E.Property(cell, nameof(ICell.DateCellValue)), valueType);
        }

        return Type.GetTypeCode(valueType) switch {
            TypeCode.String => E.Property(cell, nameof(ICell.StringCellValue)),
            TypeCode.Boolean => E.Property(cell, nameof(ICell.BooleanCellValue)),
            TypeCode.DateTime => E.Property(cell, nameof(ICell.DateCellValue)),
            TypeCode.Double => E.Property(cell, nameof(ICell.NumericCellValue)),
            >= TypeCode.SByte and <= TypeCode.Decimal
                => E.Convert(E.Property(cell, nameof(ICell.NumericCellValue)), valueType),
            _ => throw new NotSupportedException($"Cell does not support this type({valueType}) of read."),
        };
    }
}
