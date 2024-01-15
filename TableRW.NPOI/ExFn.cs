

namespace TableRW;

internal static class ExcelEx {
    internal static V GetValueOr<K, V>(this Dictionary<K, V> dic, K key, V orValue)
    => dic.TryGetValue(key, out var val) ? val : orValue;

}
