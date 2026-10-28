/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System.Collections;
using System.Diagnostics.CodeAnalysis;

/// <summary>
/// Represents a dictionary where each key can be associated with multiple values.
/// It functions as a standard read-only dictionary, but the value associated 
/// with each key is an enumerable collection of values.
/// </summary>
sealed class MultiValueDictionary<TKey, TValue>
    (IEqualityComparer<TKey> comparer)
        : IReadOnlyDictionary<TKey, IReadOnlyCollection<TValue>>
          where TKey: notnull
{
    private readonly Dictionary<TKey, List<TValue>> dictionary = new(comparer);

    public IEnumerable<TKey> Keys => dictionary.Keys;

    public IEnumerable<IReadOnlyCollection<TValue>> Values => dictionary.Values;

    public int Count => dictionary.Count;

    /// <summary>
    /// Adds a value associated with the given key. 
    /// If the key exists, the value is added to its collection; otherwise, a new association is created.
    /// </summary>
    public void AddWithValue(TKey key, TValue value)
    {
        if (!dictionary.TryGetValue(key, out var values))
        {
            values = [];
            dictionary[key] = values;
        }
        values.Add(value);
    }

    /// <summary>
    /// Attempts to retrieve the collection of values associated with the specified key.
    /// </summary>
    public bool TryGetValue(TKey key, [NotNullWhen(true)] out IReadOnlyCollection<TValue> value)
    {
        var success = dictionary.TryGetValue(key, out var collection);
        value = collection!;
        return success;
    }

    /// <summary>
    /// Retrieves the most recently added value associated with the specified key.
    /// </summary>
    public bool TryGetSingleValue(TKey key, [MaybeNullWhen(false)] out TValue value)
    {
        var success = dictionary.TryGetValue(key, out var collection);
        if (success) value = collection![^1];
        else value = default;

        return success;
    }

    /// <summary>
    /// Determines if the given key exists and has at least one value associated with it.
    /// </summary>
    public bool ContainsKey(TKey key)
    {
        return dictionary.ContainsKey(key);
    }

    public IReadOnlyCollection<TValue> this[TKey key] 
    {
        get
        {
            if (key == null)
                throw new ArgumentNullException(nameof(key));

            if (!dictionary.TryGetValue(key, out var collection))
                throw new KeyNotFoundException();
            return collection;

        }
    }

    public IEnumerator<KeyValuePair<TKey, IReadOnlyCollection<TValue>>> GetEnumerator()
        => ((IReadOnlyDictionary<TKey, IReadOnlyCollection<TValue>>) dictionary).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator()
        => dictionary.GetEnumerator();
}