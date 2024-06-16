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
using System;
using System.Collections;
using System.Collections.Generic;

namespace HtmlToOpenXml;

/// <summary>
/// Collection which contains the remaining row span per column.
/// </summary>
sealed class RowSpanCollection : IEnumerable<KeyValuePair<int, RowSpanCollection.CellSpan>>
{
#if NET5_0_OR_GREATER
    public readonly record struct CellSpan(int RowSpan, int Colspan);
#else
    public readonly struct CellSpan(int rowSpan, int colSpan)
    {
        public readonly int RowSpan = rowSpan;
        public readonly int Colspan = colSpan;
    }
#endif

    /// <summary>Hold the remaining row span value per colum index</summary>
    private readonly SortedDictionary<int, CellSpan> spans = [];

    /// <summary>
    /// Register a new row span.
    /// </summary>
    /// <param name="index">The index of the column</param>
    /// <param name="rowSpan">Indicates for how many rows the data cell spans or extends</param>
    /// <param name="columnSpan">Whether the row span must be carried on the next columns</param>
    public void Add(int index, int rowSpan, int columnSpan = 1)
    {
        if (rowSpan == 0) rowSpan = int.MaxValue;
        else rowSpan--;

        spans.Add(index, new CellSpan(rowSpan, Math.Max(1, columnSpan)));
    }

    /// <summary>
    /// Carry on the rown span on the next row.
    /// </summary>
    public int Decrement(int columnIndex)
    {
        if (!spans.TryGetValue(columnIndex, out var val))
            return 1;

        if (val.RowSpan <= 1) spans.Remove(columnIndex);
        else spans[columnIndex] = new CellSpan(val.RowSpan - 1, val.Colspan);
        return val.Colspan;
    }

    /// <summary>
    /// Reconciliate the carried row span from previous rows with the current spans.
    /// </summary>
    public void UnionWith(IEnumerable<KeyValuePair<int, CellSpan>> other)
    {
        foreach (var span in other)
            spans.Add(span.Key, span.Value);
    }

    public IEnumerator<KeyValuePair<int, CellSpan>> GetEnumerator()
    {
        return spans.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    /// <summary>
    /// Iterate through the column indexes of carried row span.
    /// </summary>
    public IEnumerable<int> Columns
    {
        get => spans.Keys;
    }
}