using NPOI.SS.UserModel;
using E = System.Linq.Expressions.Expression;

namespace TableRW.Write.I.NpoiEx;
#pragma warning disable CS1587

public class SheetWriterImpl<C> : TableWriterImpl<C> {

    static SheetWriterImpl() {
        WriteSource<ISheet>.Impl(WriteSrcValue, InitRow);
    }

    public static Expression[] InitRow(Expression ctx) {
        var row = // ctx.Row = ctx.Src.CreateRow(it.iRow)
            E.Assign(
                E.Property(ctx, "Row"),
                E.Call(E.Property(ctx, "Src"), "CreateRow", [], E.Property(ctx, "iRow")));
        return [row];
    }
    public static Expression WriteSrcValue(Expression ctx, Expression value) {
        // ctx.Row.CreateCell(ctx.iCol)
        var createCell = E.Call(E.Property(ctx, "Row"), "CreateCell", [], E.Property(ctx, "iCol"));

        if (Nullable.GetUnderlyingType(value.Type) != null) {
            // null 也会创建一个 cell，让 ctx.Row.Cells[iCol] 不至于返回 null cell
            return E.IfThenElse(E.Property(value, "HasValue"),
                WriteSrcValue(ctx, E.Property(value, "Value")),
                createCell);
        }

        /// <see cref="ICell.SetCellValue"/>
        var convertVal = Type.GetTypeCode(value.Type) switch {
            TypeCode.Single
            or (>= TypeCode.SByte and <= TypeCode.UInt32)
            //or TypeCode.Decimal // Loss of precision, no conversion
                => E.Convert(value, typeof(double)),
            TypeCode.Char or
            TypeCode.Int64 or TypeCode.UInt64 // Converting to double type loses precision, use string type
                => E.Call(value, "ToString", []),
            TypeCode.String or TypeCode.Boolean or
            TypeCode.Double or TypeCode.DateTime => value,
            _ => throw new NotSupportedException($"The type({value.Type}) of value does not determine the writing method"),
        };
        return E.Call(createCell, "SetCellValue", [], convertVal);
    }

}
