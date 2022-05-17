using System.Collections.Generic;
using System.Linq;

namespace ISMExtensions.Extensions
{
    public static class DictionaryExtensions
    {
        public static int? GetIndexByKey(this Dictionary<string,int> data , string key)
        {
            if (data.ContainsKey(key))
                return data[key];
            return
                null;
        }

        public static IEnumerable<V> GetValues<K, V>(this IDictionary<K, V> dict, IEnumerable<K> keys)
        {
            return keys.Select((x) => dict[x]);
        }
    }
}
