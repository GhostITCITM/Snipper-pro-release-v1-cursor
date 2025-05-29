using System.Collections.Generic;

namespace SnipperClone.Core
{
    public static class DictionaryExtensions
    {
        public static TValue GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue = default)
        {
            if (dictionary == null)
                return defaultValue;
            return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
        }
    }
} 