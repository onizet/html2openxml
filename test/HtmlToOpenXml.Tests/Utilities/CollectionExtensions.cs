// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

namespace System.Linq;

#if NET462

/// <summary>
/// Helper class that provide some extension methods to <see cref="IEnumerable{T}"/> API.
/// </summary>
static class CollectionExtensions
{
    /// <summary>
    /// Attempts to add the specified key and value to the dictionary.
    /// </summary>
    /// <param name="dictionary">The dictionary in which to insert the item.</param>
    /// <param name="key">The key of the element to add.</param>
    /// <param name="value">The value of the element to add. It can be <see langword="null"/>.</param>
    /// <returns><see langword="true"/> if the key/value pair was added to the dictionary successfully; otherwise, <see langword="false"/>.</returns>
    /// <summary>Returns the maximum value in a generic sequence according to a specified key selector function.</summary>
    /// <typeparam name="TSource">The type of the elements of <paramref name="source" />.</typeparam>
    /// <typeparam name="TKey">The type of key to compare elements by.</typeparam>
    /// <param name="source">A sequence of values to determine the maximum value of.</param>
    /// <param name="keySelector">A function to extract the key for each element.</param>
    /// <param name="comparer">The <see cref="IComparer{TKey}" /> to compare keys.</param>
    /// <returns>The value with the maximum key in the sequence.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="source" /> is <see langword="null" />.</exception>
    /// <exception cref="ArgumentException">No key extracted from <paramref name="source" /> implements the <see cref="IComparable" /> or <see cref="IComparable{TKey}" /> interface.</exception>
    /// <remarks>
    /// <para>If <typeparamref name="TKey" /> is a reference type and the source sequence is empty or contains only values that are <see langword="null" />, this method returns <see langword="null" />.</para>
    /// </remarks>
    public static TSource? MaxBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));

        if (keySelector == null)
            throw new ArgumentNullException(nameof(keySelector));

        var comparer = Comparer<TKey>.Default;

        using IEnumerator<TSource> e = source.GetEnumerator();

        if (!e.MoveNext())
        {
            if (default(TSource) is null)
            {
                return default;
            }
            else
            {
                throw new InvalidOperationException("Sequence contains no elements");
            }
        }

        TSource value = e.Current;
        TKey key = keySelector(value);

        if (default(TKey) is null)
        {
            if (key == null)
            {
                TSource firstValue = value;

                do
                {
                    if (!e.MoveNext())
                    {
                        // All keys are null, surface the first element.
                        return firstValue;
                    }

                    value = e.Current;
                    key = keySelector(value);
                }
                while (key == null);
            }

            while (e.MoveNext())
            {
                TSource nextValue = e.Current;
                TKey nextKey = keySelector(nextValue);
                if (nextKey != null && comparer.Compare(nextKey, key) > 0)
                {
                    key = nextKey;
                    value = nextValue;
                }
            }
        }
        else
        {
            if (comparer == Comparer<TKey>.Default)
            {
                while (e.MoveNext())
                {
                    TSource nextValue = e.Current;
                    TKey nextKey = keySelector(nextValue);
                    if (Comparer<TKey>.Default.Compare(nextKey, key) > 0)
                    {
                        key = nextKey;
                        value = nextValue;
                    }
                }
            }
            else
            {
                while (e.MoveNext())
                {
                    TSource nextValue = e.Current;
                    TKey nextKey = keySelector(nextValue);
                    if (comparer.Compare(nextKey, key) > 0)
                    {
                        key = nextKey;
                        value = nextValue;
                    }
                }
            }
        }

        return value;
    }
}
#endif
